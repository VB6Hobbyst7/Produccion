VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPagoServRecaudo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servicios - Servicios de Recaudo"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8445
   Icon            =   "frmPagoServRecaudo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Selección del Convenio"
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
      Height          =   10260
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   8415
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
         Height          =   375
         Left            =   5550
         TabIndex        =   40
         Top             =   9780
         Width           =   1335
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
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
         TabIndex        =   41
         Top             =   9720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalir 
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
         Height          =   375
         Left            =   150
         TabIndex        =   42
         Top             =   9780
         Width           =   1335
      End
      Begin TabDlg.SSTab stContenedorValidacion 
         Height          =   7920
         Left            =   105
         TabIndex        =   48
         Top             =   1800
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   13970
         _Version        =   393216
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Validación Completa"
         TabPicture(0)   =   "frmPagoServRecaudo.frx":030A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label6"
         Tab(0).Control(1)=   "Label7"
         Tab(0).Control(2)=   "Label8"
         Tab(0).Control(3)=   "Frame3"
         Tab(0).Control(4)=   "Frame4"
         Tab(0).Control(5)=   "grdConceptoPagarVC"
         Tab(0).Control(6)=   "txtSubTotalVC"
         Tab(0).Control(7)=   "txtComisionVC"
         Tab(0).Control(8)=   "txtTotalVC"
         Tab(0).Control(9)=   "fraFormaPago"
         Tab(0).Control(10)=   "fraTranfereciaVC"
         Tab(0).ControlCount=   11
         TabCaption(1)   =   "Validación Incompleta"
         TabPicture(1)   =   "frmPagoServRecaudo.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label13"
         Tab(1).Control(1)=   "Label14"
         Tab(1).Control(2)=   "Label15"
         Tab(1).Control(3)=   "grdConceptoPagarVI"
         Tab(1).Control(4)=   "Frame5"
         Tab(1).Control(5)=   "Frame6"
         Tab(1).Control(6)=   "txtSubTotalVI"
         Tab(1).Control(7)=   "txtComisionVI"
         Tab(1).Control(8)=   "txtTotalVI"
         Tab(1).Control(9)=   "Frame2"
         Tab(1).Control(10)=   "fraTranfereciaVI"
         Tab(1).ControlCount=   11
         TabCaption(2)   =   "Sin Validación"
         TabPicture(2)   =   "frmPagoServRecaudo.frx":0342
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label19"
         Tab(2).Control(1)=   "Label20"
         Tab(2).Control(2)=   "Label21"
         Tab(2).Control(3)=   "grdConceptoPagarSV"
         Tab(2).Control(4)=   "Frame8"
         Tab(2).Control(5)=   "txtSubTotalSV"
         Tab(2).Control(6)=   "txtComisionSV"
         Tab(2).Control(7)=   "txtTotalSV"
         Tab(2).Control(8)=   "cmdAgregarSV"
         Tab(2).Control(9)=   "cmdQuitarSV"
         Tab(2).Control(10)=   "Frame9"
         Tab(2).Control(11)=   "fraTranfereciaSV"
         Tab(2).ControlCount=   12
         TabCaption(3)   =   "Validación por Importe"
         TabPicture(3)   =   "frmPagoServRecaudo.frx":035E
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "Label29"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "Label30"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "Label31"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "grdConceptoPagarVX"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).Control(4)=   "txtSubTotalVX"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).Control(5)=   "txtComisionVX"
         Tab(3).Control(5).Enabled=   0   'False
         Tab(3).Control(6)=   "txtTotalVX"
         Tab(3).Control(6).Enabled=   0   'False
         Tab(3).Control(7)=   "cmdAgregarVX"
         Tab(3).Control(7).Enabled=   0   'False
         Tab(3).Control(8)=   "cmdQuitarVX"
         Tab(3).Control(8).Enabled=   0   'False
         Tab(3).Control(9)=   "Frame7"
         Tab(3).Control(9).Enabled=   0   'False
         Tab(3).Control(10)=   "Frame10"
         Tab(3).Control(10).Enabled=   0   'False
         Tab(3).Control(11)=   "fraTranfereciaVP"
         Tab(3).Control(11).Enabled=   0   'False
         Tab(3).ControlCount=   12
         Begin VB.Frame fraTranfereciaVP 
            Caption         =   "Transferencia"
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
            Height          =   1725
            Left            =   120
            TabIndex        =   160
            Top             =   2640
            Width           =   7935
            Begin VB.TextBox txtTransferGlosaVP 
               Appearance      =   0  'Flat
               Height          =   555
               Left            =   840
               MaxLength       =   255
               TabIndex        =   163
               Top             =   1050
               Width           =   2865
            End
            Begin VB.CommandButton cmdTranferVP 
               Height          =   350
               Left            =   3240
               Picture         =   "frmPagoServRecaudo.frx":037A
               Style           =   1  'Graphical
               TabIndex        =   162
               Top             =   645
               Width           =   475
            End
            Begin VB.ComboBox cboTransferMonedaVP 
               Height          =   315
               Left            =   855
               Style           =   2  'Dropdown List
               TabIndex        =   161
               Top             =   255
               Width           =   2880
            End
            Begin VB.Label Label67 
               Caption         =   "TCV"
               Height          =   285
               Left            =   6555
               TabIndex        =   175
               Top             =   270
               Width           =   390
            End
            Begin VB.Label Label66 
               Caption         =   "TCC"
               Height          =   285
               Left            =   4050
               TabIndex        =   174
               Top             =   270
               Width           =   390
            End
            Begin VB.Label Label65 
               AutoSize        =   -1  'True
               Caption         =   "Glosa :"
               Height          =   195
               Left            =   120
               TabIndex        =   173
               Top             =   1170
               Width           =   495
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Left            =   45
               TabIndex        =   172
               Top             =   315
               Width           =   585
            End
            Begin VB.Label lblTrasferNDVP 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   840
               TabIndex        =   171
               Top             =   645
               Width           =   2280
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               Caption         =   "Nro Doc :"
               Height          =   195
               Left            =   45
               TabIndex        =   170
               Top             =   720
               Width           =   690
            End
            Begin VB.Label lbltransferBcoVP 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   3840
               TabIndex        =   169
               Top             =   645
               Width           =   3945
            End
            Begin VB.Label lblTTCCDVP 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4545
               TabIndex        =   168
               Top             =   255
               Width           =   750
            End
            Begin VB.Label lblTTCVDVP 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   7050
               TabIndex        =   167
               Top             =   240
               Width           =   750
            End
            Begin VB.Label lblEtiMonTraVP 
               AutoSize        =   -1  'True
               Caption         =   "Monto Transacción"
               Height          =   195
               Left            =   3840
               TabIndex        =   166
               Top             =   1140
               Visible         =   0   'False
               Width           =   1380
            End
            Begin VB.Label lblSimTraVP 
               AutoSize        =   -1  'True
               Caption         =   "S/."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   240
               Left            =   5400
               TabIndex        =   165
               Top             =   1110
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.Label lblMonTraVP 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
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
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   6120
               TabIndex        =   164
               Top             =   1080
               Visible         =   0   'False
               Width           =   1665
            End
         End
         Begin VB.Frame fraTranfereciaSV 
            Caption         =   "Transferencia"
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
            Height          =   1725
            Left            =   -74880
            TabIndex        =   144
            Top             =   2640
            Width           =   7935
            Begin VB.ComboBox cboTransferMonedaSV 
               Height          =   315
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   147
               Top             =   240
               Width           =   2880
            End
            Begin VB.CommandButton cmdTranferSV 
               Height          =   350
               Left            =   3240
               Picture         =   "frmPagoServRecaudo.frx":0684
               Style           =   1  'Graphical
               TabIndex        =   146
               Top             =   645
               Width           =   475
            End
            Begin VB.TextBox txtTransferGlosaSV 
               Appearance      =   0  'Flat
               Height          =   555
               Left            =   855
               MaxLength       =   255
               TabIndex        =   145
               Top             =   1080
               Width           =   2865
            End
            Begin VB.Label lblMonTraSV 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
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
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   6120
               TabIndex        =   159
               Top             =   1080
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.Label lblSimTraSV 
               AutoSize        =   -1  'True
               Caption         =   "S/."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   240
               Left            =   5400
               TabIndex        =   158
               Top             =   1110
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.Label lblEtiMonTraSV 
               AutoSize        =   -1  'True
               Caption         =   "Monto Transacción"
               Height          =   195
               Left            =   3840
               TabIndex        =   157
               Top             =   1140
               Visible         =   0   'False
               Width           =   1380
            End
            Begin VB.Label lblTTCVDSV 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   7050
               TabIndex        =   156
               Top             =   240
               Width           =   750
            End
            Begin VB.Label lblTTCCDSV 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4545
               TabIndex        =   155
               Top             =   255
               Width           =   750
            End
            Begin VB.Label lbltransferBcoSV 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   3840
               TabIndex        =   154
               Top             =   645
               Width           =   3945
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               Caption         =   "Nro Doc :"
               Height          =   195
               Left            =   45
               TabIndex        =   153
               Top             =   720
               Width           =   690
            End
            Begin VB.Label lblTrasferNDSV 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   840
               TabIndex        =   152
               Top             =   645
               Width           =   2280
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Left            =   45
               TabIndex        =   151
               Top             =   315
               Width           =   585
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               Caption         =   "Glosa :"
               Height          =   195
               Left            =   120
               TabIndex        =   150
               Top             =   1170
               Width           =   495
            End
            Begin VB.Label Label40 
               Caption         =   "TCC"
               Height          =   285
               Left            =   4050
               TabIndex        =   149
               Top             =   270
               Width           =   390
            End
            Begin VB.Label Label39 
               Caption         =   "TCV"
               Height          =   285
               Left            =   6555
               TabIndex        =   148
               Top             =   270
               Width           =   390
            End
         End
         Begin VB.Frame fraTranfereciaVI 
            Caption         =   "Transferencia"
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
            Height          =   1725
            Left            =   -74880
            TabIndex        =   128
            Top             =   2760
            Width           =   7935
            Begin VB.TextBox txtTransferGlosaVI 
               Appearance      =   0  'Flat
               Height          =   555
               Left            =   855
               MaxLength       =   255
               TabIndex        =   131
               Top             =   1050
               Width           =   2865
            End
            Begin VB.CommandButton cmdTranferVI 
               Height          =   350
               Left            =   3240
               Picture         =   "frmPagoServRecaudo.frx":098E
               Style           =   1  'Graphical
               TabIndex        =   130
               Top             =   645
               Width           =   475
            End
            Begin VB.ComboBox cboTransferMonedaVI 
               Height          =   315
               Left            =   855
               Style           =   2  'Dropdown List
               TabIndex        =   129
               Top             =   255
               Width           =   2880
            End
            Begin VB.Label Label50 
               Caption         =   "TCV"
               Height          =   285
               Left            =   6555
               TabIndex        =   143
               Top             =   270
               Width           =   390
            End
            Begin VB.Label Label49 
               Caption         =   "TCC"
               Height          =   285
               Left            =   4050
               TabIndex        =   142
               Top             =   240
               Width           =   390
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               Caption         =   "Glosa :"
               Height          =   195
               Left            =   120
               TabIndex        =   141
               Top             =   1170
               Width           =   495
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Left            =   45
               TabIndex        =   140
               Top             =   315
               Width           =   585
            End
            Begin VB.Label lblTrasferNDVI 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   840
               TabIndex        =   139
               Top             =   645
               Width           =   2280
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               Caption         =   "Nro Doc :"
               Height          =   195
               Left            =   45
               TabIndex        =   138
               Top             =   720
               Width           =   690
            End
            Begin VB.Label lbltransferBcoVI 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   3840
               TabIndex        =   137
               Top             =   645
               Width           =   3945
            End
            Begin VB.Label lblTTCCDVI 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4545
               TabIndex        =   136
               Top             =   255
               Width           =   750
            End
            Begin VB.Label lblTTCVDVI 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   7050
               TabIndex        =   135
               Top             =   240
               Width           =   750
            End
            Begin VB.Label lblEtiMonTraVI 
               AutoSize        =   -1  'True
               Caption         =   "Monto Transacción"
               Height          =   195
               Left            =   3840
               TabIndex        =   134
               Top             =   1140
               Visible         =   0   'False
               Width           =   1380
            End
            Begin VB.Label lblSimTraVI 
               AutoSize        =   -1  'True
               Caption         =   "S/."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   240
               Left            =   5400
               TabIndex        =   133
               Top             =   1110
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.Label lblMonTraVI 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
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
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   6120
               TabIndex        =   132
               Top             =   1080
               Visible         =   0   'False
               Width           =   1665
            End
         End
         Begin VB.Frame fraTranfereciaVC 
            Caption         =   "Transferencia"
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
            Height          =   1725
            Left            =   -74880
            TabIndex        =   112
            Top             =   2760
            Width           =   7935
            Begin VB.ComboBox cboTransferMonedaVC 
               Height          =   315
               Left            =   855
               Style           =   2  'Dropdown List
               TabIndex        =   115
               Top             =   255
               Width           =   2880
            End
            Begin VB.CommandButton cmdTranferVC 
               Height          =   350
               Left            =   3240
               Picture         =   "frmPagoServRecaudo.frx":0C98
               Style           =   1  'Graphical
               TabIndex        =   114
               Top             =   645
               Width           =   475
            End
            Begin VB.TextBox txtTransferGlosaVC 
               Appearance      =   0  'Flat
               Height          =   555
               Left            =   855
               MaxLength       =   255
               TabIndex        =   113
               Top             =   1080
               Width           =   2865
            End
            Begin VB.Label lblMonTraVC 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
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
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   6120
               TabIndex        =   127
               Top             =   1080
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.Label lblSimTraVC 
               AutoSize        =   -1  'True
               Caption         =   "S/."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   240
               Left            =   5400
               TabIndex        =   126
               Top             =   1110
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.Label lblEtiMonTraVC 
               AutoSize        =   -1  'True
               Caption         =   "Monto Transacción"
               Height          =   195
               Left            =   3840
               TabIndex        =   125
               Top             =   1140
               Visible         =   0   'False
               Width           =   1380
            End
            Begin VB.Label lblTTCVDVC 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   7050
               TabIndex        =   124
               Top             =   255
               Width           =   750
            End
            Begin VB.Label lblTTCCDVC 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4545
               TabIndex        =   123
               Top             =   255
               Width           =   750
            End
            Begin VB.Label lbltransferBcoVC 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   3840
               TabIndex        =   122
               Top             =   645
               Width           =   3945
            End
            Begin VB.Label lbltransferN 
               AutoSize        =   -1  'True
               Caption         =   "Nro Doc :"
               Height          =   195
               Left            =   45
               TabIndex        =   121
               Top             =   720
               Width           =   690
            End
            Begin VB.Label lblTrasferNDVC 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   855
               TabIndex        =   120
               Top             =   645
               Width           =   2280
            End
            Begin VB.Label lblTransferMoneda 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               Height          =   195
               Left            =   45
               TabIndex        =   119
               Top             =   315
               Width           =   585
            End
            Begin VB.Label lblTransferGlosa 
               AutoSize        =   -1  'True
               Caption         =   "Glosa :"
               Height          =   195
               Left            =   120
               TabIndex        =   118
               Top             =   1170
               Width           =   495
            End
            Begin VB.Label lblTTCC 
               Caption         =   "TCC"
               Height          =   285
               Left            =   4050
               TabIndex        =   117
               Top             =   270
               Width           =   390
            End
            Begin VB.Label Label37 
               Caption         =   "TCV"
               Height          =   285
               Left            =   6555
               TabIndex        =   116
               Top             =   270
               Width           =   390
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Tipo de Pagp"
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
            Height          =   750
            Left            =   120
            TabIndex        =   106
            Top             =   1800
            Width           =   7935
            Begin VB.ComboBox CmbForPagVP 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   107
               Top             =   315
               Width           =   2500
            End
            Begin SICMACT.ActXCodCta txtCuentaCargoVP 
               Height          =   375
               Left            =   4200
               TabIndex        =   108
               Top             =   315
               Visible         =   0   'False
               Width           =   3630
               _ExtentX        =   6403
               _ExtentY        =   661
               Texto           =   "Cuenta N°:"
               EnabledCta      =   -1  'True
               EnabledAge      =   -1  'True
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               Caption         =   "Forma Pago"
               Height          =   195
               Left            =   180
               TabIndex        =   111
               Top             =   375
               Width           =   855
            End
            Begin VB.Label LblNumDocVP 
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
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   6105
               TabIndex        =   110
               Top             =   315
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "Nº Documento"
               Height          =   195
               Left            =   5025
               TabIndex        =   109
               Top             =   375
               Visible         =   0   'False
               Width           =   1050
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Tipo de Pagp"
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
            Height          =   750
            Left            =   -74880
            TabIndex        =   100
            Top             =   1800
            Width           =   7935
            Begin VB.ComboBox CmbForPagSV 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   101
               Top             =   315
               Width           =   2500
            End
            Begin SICMACT.ActXCodCta txtCuentaCargoSV 
               Height          =   375
               Left            =   4200
               TabIndex        =   102
               Top             =   315
               Visible         =   0   'False
               Width           =   3630
               _ExtentX        =   6403
               _ExtentY        =   661
               Texto           =   "Cuenta N°:"
               EnabledCta      =   -1  'True
               EnabledAge      =   -1  'True
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               Caption         =   "Nº Documento"
               Height          =   195
               Left            =   5025
               TabIndex        =   105
               Top             =   375
               Visible         =   0   'False
               Width           =   1050
            End
            Begin VB.Label LblNumDocSV 
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
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   6105
               TabIndex        =   104
               Top             =   315
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Forma Pago"
               Height          =   195
               Left            =   180
               TabIndex        =   103
               Top             =   375
               Width           =   855
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Tipo de Pagp"
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
            Height          =   750
            Left            =   -74880
            TabIndex        =   94
            Top             =   1920
            Width           =   7935
            Begin VB.ComboBox CmbForPagVI 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   95
               Top             =   315
               Width           =   2500
            End
            Begin SICMACT.ActXCodCta txtCuentaCargoVI 
               Height          =   375
               Left            =   4200
               TabIndex        =   96
               Top             =   315
               Visible         =   0   'False
               Width           =   3630
               _ExtentX        =   6403
               _ExtentY        =   661
               Texto           =   "Cuenta N°:"
               EnabledCta      =   -1  'True
               EnabledAge      =   -1  'True
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Forma Pago"
               Height          =   195
               Left            =   180
               TabIndex        =   99
               Top             =   375
               Width           =   855
            End
            Begin VB.Label LblNumDocVI 
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
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   6105
               TabIndex        =   98
               Top             =   315
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Nº Documento"
               Height          =   195
               Left            =   5025
               TabIndex        =   97
               Top             =   375
               Visible         =   0   'False
               Width           =   1050
            End
         End
         Begin VB.Frame fraFormaPago 
            Caption         =   "Tipo de Pagp"
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
            Height          =   750
            Left            =   -74880
            TabIndex        =   88
            Top             =   1920
            Width           =   7935
            Begin VB.ComboBox CmbForPagVC 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   89
               Top             =   315
               Width           =   2500
            End
            Begin SICMACT.ActXCodCta txtCuentaCargoVC 
               Height          =   375
               Left            =   4200
               TabIndex        =   90
               Top             =   315
               Visible         =   0   'False
               Width           =   3630
               _ExtentX        =   6403
               _ExtentY        =   661
               Texto           =   "Cuenta N°:"
               EnabledCta      =   -1  'True
               EnabledAge      =   -1  'True
            End
            Begin VB.Label lblNroDocumento 
               AutoSize        =   -1  'True
               Caption         =   "Nº Documento"
               Height          =   195
               Left            =   5025
               TabIndex        =   93
               Top             =   375
               Visible         =   0   'False
               Width           =   1050
            End
            Begin VB.Label LblNumDocVC 
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
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   6105
               TabIndex        =   92
               Top             =   315
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.Label lblFormaPago 
               AutoSize        =   -1  'True
               Caption         =   "Forma Pago"
               Height          =   195
               Left            =   180
               TabIndex        =   91
               Top             =   375
               Width           =   855
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Operación"
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
            Height          =   1425
            Left            =   105
            TabIndex        =   78
            Top             =   360
            Width           =   7935
            Begin VB.TextBox txtDOIVX 
               Appearance      =   0  'Flat
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
               Left            =   3810
               TabIndex        =   32
               Top             =   960
               Width           =   1380
            End
            Begin VB.TextBox txtNombreClienteVX 
               Appearance      =   0  'Flat
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
               Left            =   1065
               TabIndex        =   30
               Top             =   600
               Width           =   6390
            End
            Begin VB.ComboBox cboTipoDOIVX 
               Height          =   315
               ItemData        =   "frmPagoServRecaudo.frx":0FA2
               Left            =   1065
               List            =   "frmPagoServRecaudo.frx":0FAC
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox txtOtroCodigoVX 
               Appearance      =   0  'Flat
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
               Left            =   5970
               MaxLength       =   10
               TabIndex        =   33
               Top             =   960
               Width           =   1515
            End
            Begin VB.Label Label33 
               Caption         =   "Nombre: "
               Height          =   255
               Left            =   150
               TabIndex        =   84
               Top             =   660
               Width           =   735
            End
            Begin VB.Label Label28 
               Caption         =   "Tipo DOI: "
               Height          =   255
               Left            =   150
               TabIndex        =   83
               Top             =   1020
               Width           =   945
            End
            Begin VB.Label Label27 
               Caption         =   "Nª DOI: "
               Height          =   255
               Left            =   2520
               TabIndex        =   82
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label26 
               Caption         =   "Código: "
               Height          =   255
               Left            =   5280
               TabIndex        =   81
               Top             =   1020
               Width           =   615
            End
            Begin VB.Label lblDescripcionCodigoVP 
               Alignment       =   2  'Center
               Caption         =   "Descripción de la Operación"
               Height          =   255
               Left            =   1065
               TabIndex        =   80
               Top             =   300
               Width           =   6390
            End
            Begin VB.Label Label18 
               Caption         =   "Nª DOI:"
               Height          =   255
               Left            =   3030
               TabIndex        =   79
               Top             =   990
               Width           =   735
            End
         End
         Begin VB.CommandButton cmdQuitarVX 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   1260
            TabIndex        =   36
            Top             =   6840
            Width           =   1095
         End
         Begin VB.CommandButton cmdAgregarVX 
            Caption         =   "&Agregar"
            Height          =   375
            Left            =   90
            TabIndex        =   35
            Top             =   6840
            Width           =   1095
         End
         Begin VB.TextBox txtTotalVX 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
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
            Left            =   6450
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "0.00"
            Top             =   7440
            Width           =   1575
         End
         Begin VB.TextBox txtComisionVX 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   6450
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "0.00"
            Top             =   7125
            Width           =   1575
         End
         Begin VB.TextBox txtSubTotalVX 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   6450
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "0.00"
            Top             =   6810
            Width           =   1575
         End
         Begin VB.CommandButton cmdQuitarSV 
            Caption         =   "Quitar"
            Height          =   375
            Left            =   -73680
            TabIndex        =   26
            Top             =   6915
            Width           =   1095
         End
         Begin VB.CommandButton cmdAgregarSV 
            Caption         =   "&Agregar"
            Height          =   375
            Left            =   -74880
            TabIndex        =   25
            Top             =   6915
            Width           =   1095
         End
         Begin VB.TextBox txtTotalSV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
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
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "0.00"
            Top             =   7440
            Width           =   1575
         End
         Begin VB.TextBox txtComisionSV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "0.00"
            Top             =   7125
            Width           =   1575
         End
         Begin VB.TextBox txtSubTotalSV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "0.00"
            Top             =   6810
            Width           =   1575
         End
         Begin VB.Frame Frame8 
            Caption         =   "Operación"
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
            Height          =   1425
            Left            =   -74880
            TabIndex        =   65
            Top             =   360
            Width           =   7935
            Begin VB.TextBox txtDOISV 
               Appearance      =   0  'Flat
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
               Left            =   3690
               TabIndex        =   22
               ToolTipText     =   "Ingresar DOI"
               Top             =   960
               Width           =   1380
            End
            Begin VB.TextBox txtOtroCodigoSV 
               Appearance      =   0  'Flat
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
               Left            =   5970
               MaxLength       =   10
               TabIndex        =   23
               ToolTipText     =   "Ingresar Codigo"
               Top             =   960
               Width           =   1485
            End
            Begin VB.ComboBox cboTipoDOISV 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "frmPagoServRecaudo.frx":120D
               Left            =   1065
               List            =   "frmPagoServRecaudo.frx":1217
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox txtNombreClienteSV 
               Appearance      =   0  'Flat
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
               Left            =   1065
               TabIndex        =   20
               Top             =   600
               Width           =   6390
            End
            Begin VB.Label Label32 
               Caption         =   "Nª DOI:"
               Height          =   255
               Left            =   3030
               TabIndex        =   77
               Top             =   1020
               Width           =   735
            End
            Begin VB.Label lblDescripcionCodigoSV 
               Alignment       =   2  'Center
               Caption         =   "Descripción de la Operación"
               Height          =   255
               Left            =   1050
               TabIndex        =   76
               Top             =   300
               Width           =   6405
            End
            Begin VB.Label Label24 
               Caption         =   "Código: "
               Height          =   255
               Left            =   5190
               TabIndex        =   72
               Top             =   1020
               Width           =   735
            End
            Begin VB.Label Label23 
               Caption         =   "Nª DOI: "
               Height          =   255
               Left            =   2520
               TabIndex        =   71
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label22 
               Caption         =   "Tipo DOI: "
               Height          =   255
               Left            =   150
               TabIndex        =   70
               Top             =   1020
               Width           =   750
            End
            Begin VB.Label Label17 
               Caption         =   "Nombre: "
               Height          =   255
               Left            =   180
               TabIndex        =   66
               Top             =   630
               Width           =   735
            End
         End
         Begin VB.TextBox txtTotalVI 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
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
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "0.00"
            Top             =   7455
            Width           =   1575
         End
         Begin VB.TextBox txtComisionVI 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "0.00"
            Top             =   7140
            Width           =   1575
         End
         Begin VB.TextBox txtSubTotalVI 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "0.00"
            Top             =   6825
            Width           =   1575
         End
         Begin VB.Frame Frame6 
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
            ForeColor       =   &H8000000D&
            Height          =   735
            Left            =   -74880
            TabIndex        =   58
            Top             =   1200
            Width           =   7935
            Begin VB.TextBox txtNombreClienteVI 
               Appearance      =   0  'Flat
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
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   270
               Width           =   4575
            End
            Begin VB.TextBox txtDOIVI 
               Appearance      =   0  'Flat
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
               Left            =   6090
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   270
               Width           =   1695
            End
            Begin VB.Label Label12 
               Caption         =   "Nombre: "
               Height          =   255
               Left            =   150
               TabIndex        =   60
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label11 
               Caption         =   "DOI: "
               Height          =   255
               Left            =   5640
               TabIndex        =   59
               Top             =   330
               Width           =   375
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Busqueda"
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
            Height          =   735
            Left            =   -74880
            TabIndex        =   57
            Top             =   360
            Width           =   7935
            Begin VB.TextBox txtCodigoIDVI 
               Appearance      =   0  'Flat
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
               Left            =   960
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   13
               Top             =   240
               Width           =   1575
            End
            Begin VB.CommandButton cmdBuscarPersonaVI 
               Appearance      =   0  'Flat
               Caption         =   "..."
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
               Left            =   2535
               TabIndex        =   14
               Top             =   240
               Width           =   400
            End
            Begin VB.Label Label34 
               Caption         =   "Código: "
               Height          =   255
               Left            =   150
               TabIndex        =   87
               Top             =   285
               Width           =   735
            End
            Begin VB.Label lblDescripcionCodigoVI 
               Caption         =   "DESCRIPCION PARA EL USUARIO"
               Height          =   285
               Left            =   3210
               TabIndex        =   86
               Top             =   300
               Width           =   4575
            End
         End
         Begin VB.TextBox txtTotalVC 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
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
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0.00"
            Top             =   7455
            Width           =   1575
         End
         Begin VB.TextBox txtComisionVC 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   7140
            Width           =   1575
         End
         Begin VB.TextBox txtSubTotalVC 
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
            Left            =   -68520
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   6825
            Width           =   1575
         End
         Begin SICMACT.FlexEdit grdConceptoPagarVC 
            Height          =   2235
            Left            =   -74880
            TabIndex        =   9
            Top             =   4560
            Width           =   7935
            _extentx        =   13996
            _extenty        =   2355
            cols0           =   10
            fixedcols       =   0
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "ID-Servicio-Concepto-Fec Vec-Moneda-Importe-Mora-Pagar-IDCopy-a"
            encabezadosanchos=   "0-1900-1900-0-1200-1300-0-1000-0-0"
            font            =   "frmPagoServRecaudo.frx":1376
            font            =   "frmPagoServRecaudo.frx":13A2
            font            =   "frmPagoServRecaudo.frx":13CE
            font            =   "frmPagoServRecaudo.frx":13FA
            font            =   "frmPagoServRecaudo.frx":1426
            fontfixed       =   "frmPagoServRecaudo.frx":1452
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            columnasaeditar =   "X-X-X-X-X-X-X-7-X-X"
            listacontroles  =   "0-0-0-0-0-0-0-4-0-0"
            encabezadosalineacion=   "C-L-L-C-L-R-R-C-C-C"
            formatosedit    =   "0-0-0-0-2-0-0-0-0-0"
            textarray0      =   "ID"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin VB.Frame Frame4 
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
            ForeColor       =   &H8000000D&
            Height          =   735
            Left            =   -74880
            TabIndex        =   50
            Top             =   1080
            Width           =   7935
            Begin VB.TextBox txtDOIVC 
               Appearance      =   0  'Flat
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
               Left            =   6120
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   300
               Width           =   1695
            End
            Begin VB.TextBox txtNombreClienteVC 
               Appearance      =   0  'Flat
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
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   300
               Width           =   4575
            End
            Begin VB.Label Label5 
               Caption         =   "DOI: "
               Height          =   255
               Left            =   5640
               TabIndex        =   52
               Top             =   360
               Width           =   375
            End
            Begin VB.Label Label4 
               Caption         =   "Nombre: "
               Height          =   255
               Left            =   180
               TabIndex        =   51
               Top             =   330
               Width           =   735
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Busqueda"
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
            Height          =   735
            Left            =   -74880
            TabIndex        =   49
            Top             =   360
            Width           =   7935
            Begin VB.CommandButton cmdBuscarPersonaVC 
               Appearance      =   0  'Flat
               Caption         =   "..."
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
               Left            =   2535
               TabIndex        =   6
               Top             =   270
               Width           =   400
            End
            Begin VB.TextBox txtCodigoIDVC 
               Appearance      =   0  'Flat
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
               Left            =   960
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   5
               Top             =   270
               Width           =   1575
            End
            Begin VB.Label Label9 
               Caption         =   "Código: "
               Height          =   255
               Left            =   150
               TabIndex        =   85
               Top             =   300
               Width           =   630
            End
            Begin VB.Label lblDescripcionCodigoVC 
               Caption         =   "DESCRIPCION PARA EL USUARIO"
               Height          =   225
               Left            =   3240
               TabIndex        =   56
               Top             =   330
               Width           =   4590
            End
         End
         Begin SICMACT.FlexEdit grdConceptoPagarVI 
            Height          =   2265
            Left            =   -74880
            TabIndex        =   61
            Top             =   4560
            Width           =   7935
            _extentx        =   13996
            _extenty        =   2963
            cols0           =   8
            fixedcols       =   0
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "ID-Servicio-Concepto-Moneda-Pagar-Importe-Deuda-a"
            encabezadosanchos=   "500-1900-1900-1200-1000-1300-0-0"
            font            =   "frmPagoServRecaudo.frx":1480
            font            =   "frmPagoServRecaudo.frx":14AC
            font            =   "frmPagoServRecaudo.frx":14D8
            font            =   "frmPagoServRecaudo.frx":1504
            font            =   "frmPagoServRecaudo.frx":1530
            fontfixed       =   "frmPagoServRecaudo.frx":155C
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            columnasaeditar =   "X-X-X-X-4-5-X-X"
            listacontroles  =   "0-0-0-0-4-0-0-0"
            encabezadosalineacion=   "C-C-C-L-C-R-C-C"
            formatosedit    =   "0-0-0-0-0-2-2-2"
            textarray0      =   "ID"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            colwidth0       =   495
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit grdConceptoPagarSV 
            Height          =   2400
            Left            =   -74880
            TabIndex        =   24
            Top             =   4410
            Width           =   7935
            _extentx        =   13996
            _extenty        =   2355
            cols0           =   5
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Concepto-Moneda-Importe-id"
            encabezadosanchos=   "500-5000-1200-1100-0"
            font            =   "frmPagoServRecaudo.frx":158A
            font            =   "frmPagoServRecaudo.frx":15B6
            font            =   "frmPagoServRecaudo.frx":15E2
            font            =   "frmPagoServRecaudo.frx":160E
            font            =   "frmPagoServRecaudo.frx":163A
            fontfixed       =   "frmPagoServRecaudo.frx":1666
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            columnasaeditar =   "X-1-X-3-X"
            listacontroles  =   "0-0-0-0-0"
            encabezadosalineacion=   "C-L-C-R-C"
            formatosedit    =   "0-0-0-2-2"
            textarray0      =   "#"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            colwidth0       =   495
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit grdConceptoPagarVX 
            Height          =   2340
            Left            =   120
            TabIndex        =   34
            Top             =   4470
            Width           =   7935
            _extentx        =   13996
            _extenty        =   2355
            cols0           =   5
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Concepto-Moneda-Importe-id"
            encabezadosanchos=   "500-5000-1200-1100-0"
            font            =   "frmPagoServRecaudo.frx":1694
            font            =   "frmPagoServRecaudo.frx":16C0
            font            =   "frmPagoServRecaudo.frx":16EC
            font            =   "frmPagoServRecaudo.frx":1718
            font            =   "frmPagoServRecaudo.frx":1744
            fontfixed       =   "frmPagoServRecaudo.frx":1770
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            columnasaeditar =   "X-1-X-X-X"
            listacontroles  =   "0-3-0-0-0"
            encabezadosalineacion=   "C-L-C-R-C"
            formatosedit    =   "0-0-0-2-2"
            textarray0      =   "#"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
            colwidth0       =   495
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
         Begin VB.Label Label31 
            Caption         =   "Total: "
            Height          =   255
            Left            =   5640
            TabIndex        =   75
            Top             =   7500
            Width           =   735
         End
         Begin VB.Label Label30 
            Caption         =   "Comisión: "
            Height          =   255
            Left            =   5640
            TabIndex        =   74
            Top             =   7200
            Width           =   765
         End
         Begin VB.Label Label29 
            Caption         =   "Subtotal: "
            Height          =   255
            Left            =   5640
            TabIndex        =   73
            Top             =   6870
            Width           =   735
         End
         Begin VB.Label Label21 
            Caption         =   "Total: "
            Height          =   255
            Left            =   -69360
            TabIndex        =   69
            Top             =   7440
            Width           =   735
         End
         Begin VB.Label Label20 
            Caption         =   "Comisión: "
            Height          =   255
            Left            =   -69360
            TabIndex        =   68
            Top             =   7170
            Width           =   840
         End
         Begin VB.Label Label19 
            Caption         =   "Subtotal: "
            Height          =   255
            Left            =   -69360
            TabIndex        =   67
            Top             =   6870
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "Total: "
            Height          =   225
            Left            =   -69330
            TabIndex        =   64
            Top             =   7500
            Width           =   555
         End
         Begin VB.Label Label14 
            Caption         =   "Comisión: "
            Height          =   255
            Left            =   -69330
            TabIndex        =   63
            Top             =   7200
            Width           =   765
         End
         Begin VB.Label Label13 
            Caption         =   "Subtotal: "
            Height          =   255
            Left            =   -69330
            TabIndex        =   62
            Top             =   6870
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Total: "
            Height          =   255
            Left            =   -69360
            TabIndex        =   55
            Top             =   7500
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Comisión: "
            Height          =   255
            Left            =   -69360
            TabIndex        =   54
            Top             =   7200
            Width           =   765
         End
         Begin VB.Label Label6 
            Caption         =   "Subtotal: "
            Height          =   255
            Left            =   -69360
            TabIndex        =   53
            Top             =   6870
            Width           =   705
         End
      End
      Begin VB.Frame pnlBusquedaConvenio 
         Appearance      =   0  'Flat
         Caption         =   "Búsqueda de convenio"
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
         Height          =   1395
         Left            =   120
         TabIndex        =   44
         Top             =   300
         Width           =   8175
         Begin VB.CommandButton cmdBuscarConvenio 
            Appearance      =   0  'Flat
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   314
            Left            =   3255
            TabIndex        =   1
            Top             =   240
            Width           =   400
         End
         Begin VB.TextBox txtCodigoBusConvenio 
            Appearance      =   0  'Flat
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
            Left            =   1080
            MaxLength       =   18
            TabIndex        =   0
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtNombreConvenio 
            Appearance      =   0  'Flat
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
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   565
            Width           =   5505
         End
         Begin VB.TextBox txtCodigoEmpresa 
            Appearance      =   0  'Flat
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
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   900
            Width           =   1455
         End
         Begin VB.TextBox txtNombreEmpresa 
            Appearance      =   0  'Flat
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
            Left            =   2550
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   900
            Width           =   4035
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            Caption         =   "Código: "
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   300
            Width           =   690
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Caption         =   "Convenio: "
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   46
            Top             =   630
            Width           =   840
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Caption         =   "Empresa: "
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   45
            Top             =   960
            Width           =   840
         End
      End
   End
End
Attribute VB_Name = "frmPagoServRecaudo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************************************************************************
'* NOMBRE         : "frmPagoServRecaudo"
'* DESCRIPCION    : Formulario creado para el pago de servicios de convenios segun proyecto: "Mejora del Sistema y Automatizacion de Ahorros y Servicios"
'* CREACION       : RIRO, 20121213 10:00 AM
'*********************************************************************************************************************************************************

Option Explicit

Private nConvenioSeleccionado As Integer
Private rsUsuarioRecaudo As ADODB.Recordset
Private nMoneda As Integer ' 1= Soles, 2= Extranjera
Private nComisionEmpresa As Double
Private nComisionCliente As Double
Private strCuenta As String
Private clsprevio As New previo.clsprevio 'RIRO20140906
Dim Importes() As String
Dim bFocusGrid As Boolean
Dim nCliente As String 'ADD BY PTI1 20210723
Dim nConvenio As String 'ADD BY PTI1 20210723
Dim nTrama As String 'ADD BY PTI1 20210723
Dim verificaWS As Boolean 'CTI1 TI-ERS027-2019
Dim cMonedaWS As String 'CTI1 ERS027-2019
Dim cServicioWS As String 'CTI1 ERS027-2019
Private nMontoVoucher As Currency 'CTI6 ERS0112020
Dim nMovNroRVD As Long, nMovNroRVDPend As Long 'CTI6 ERS0112020
Dim sNumTarj As String 'CTI6 ERS0112020
Dim pnMoneda As Integer 'CTI6 ERS0112020
Dim psPersCodTitularAhorroCargo As String 'CTI6 ERS0112020
Dim pnITFCargoCta As Double 'CTI6 ERS0112020
Dim pbEsMismoTitular As Boolean 'CTI6 ERS0112020
Dim pnMontoPagarCargo As Double 'CTI4 ERS0112020
Dim nRedondeoITF As Double ' BRGO 20110914
Dim loVistoElectronico As frmVistoElectronico 'CTI4 ERS0112020
Dim nRespuesta As Integer 'CTI4 ERS0112020
                                           
                                         

'CTI6 ERS0112020
'****************************VC*************************************/
Dim lnTransferSaldo As Currency 'CTI7 OPEv2
Dim fsPersCodTransfer As String 'CTI7 OPEv2
Dim fnMovNroRVD As Long 'CTI7 OPEv2
Dim lnMovNroTransfer As Long 'CTI7 OPEv2
Dim lnFormaPago As Integer 'CTI4 ERS0112020
Dim lsCtaCargo As String 'CTI4 ERS0112020
Private Sub EstadoFormaPagoVC(ByVal nFormaPago As Integer)
    LblNumDocVC.Caption = ""
    txtCuentaCargoVC.NroCuenta = ""
    Select Case nFormaPago
        Case -1
            LblNumDocVC.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargoVC.Visible = False
            cmdGuardar.Enabled = True
            fraTranfereciaVC.Enabled = False
            Call IniciarVouher
        Case gCVTipoPagoEfectivo
            txtCuentaCargoVC.Visible = False
            LblNumDocVC.Visible = False
            lblNroDocumento.Visible = False
            lblNroDocumento.Visible = False
            cmdGuardar.Enabled = True
            fraTranfereciaVC.Enabled = False
            Call IniciarVouher
        Case gCVTipoPagoCargoCta
            LblNumDocVC.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargoVC.Visible = True
            txtCuentaCargoVC.Enabled = True
            txtCuentaCargoVC.CMAC = gsCodCMAC
            txtCuentaCargoVC.Prod = Trim(Str(gCapAhorros))
            cmdGuardar.Enabled = False
            fraTranfereciaVC.Enabled = False
            Call IniciarVouher
        Case gCVTipoPagoVoucher
            txtCuentaCargoVC.Visible = False
            LblNumDocVC.Visible = False
            lblNroDocumento.Visible = False
            lblNroDocumento.Visible = False
            cmdGuardar.Enabled = True
            fraTranfereciaVC.Enabled = True
    End Select
End Sub

'*****************************************************************************************************
Private Sub cboTransferMonedaVC_Click()
    If Right(cboTransferMonedaVC, 3) = Moneda.gMonedaNacional Then
        lblSimTraVC.Caption = gcPEN_SIMBOLO 'APRI20191022 SUGERENCIA CALIDAD
        lblMonTraVC.BackColor = &HC0FFFF
    Else
        lblSimTraVC.Caption = "$"
        lblMonTraVC.BackColor = &HC0FFC0
    End If
End Sub
Private Sub cboTransferMonedaVC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl Me.cmdTranferVC
    End If
End Sub
Private Sub cmdTranferVC_Click()
    Dim lsGlosa As String
    Dim lsDoc As String
    Dim lsInstit As String
    Dim oform As frmCapRegVouDepBus
    Dim lnTipMot As Integer
    Dim lsDetalle As String

    On Error GoTo ErrTransfer
    If cboTransferMonedaVC.Text = "" Then
        MsgBox "Debe escoger la moneda de la transferencia.", vbInformation, "Aviso"
        If cboTransferMonedaVC.Visible And cboTransferMonedaVC.Enabled Then EnfocaControl cboTransferMonedaVC
        Exit Sub
    End If

    lnTipMot = 16
    
    fnMovNroRVD = 0
    Set oform = New frmCapRegVouDepBus
    SetDatosTransferenciaVC "", "", "", 0, -1, ""
    oform.iniciarFormulario Trim(Right(cboTransferMonedaVC, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer, lsDetalle
    If fnMovNroRVD = 0 Then
        Exit Sub
    End If
    SetDatosTransferenciaVC lsGlosa, lsInstit, lsDoc, lnTransferSaldo, lnMovNroTransfer, lsDetalle
    Exit Sub
ErrTransfer:
    MsgBox "Ha sucedido un error al cargar los datos de la Transferencia", vbCritical, "Aviso"
End Sub
Private Sub SetDatosTransferenciaVC(ByVal psGlosa As String, ByVal psInstit As String, ByVal psDoc As String, ByVal pnTransferSaldo As Currency, ByVal pnMovNroTransfer As Long, ByVal psDetalle As String)
    Dim oPersona As New DCOMPersona
    Dim rsPersona As New ADODB.Recordset
    Dim row As Integer
    
    txtTransferGlosaVC.Text = psGlosa
    lbltransferBcoVC.Caption = psInstit
    lblTrasferNDVC.Caption = psDoc
    

    If pnMovNroTransfer <> -1 Then
        EnfocaControl txtTransferGlosaVC
    End If

    txtTransferGlosaVC.Locked = True
    lblMonTraVC = Format(pnTransferSaldo, "#,##0.00")
End Sub
'VI*********************************************************************************************
Private Sub cboTransferMonedaVI_Click()
    If Right(cboTransferMonedaVI, 3) = Moneda.gMonedaNacional Then
        lblSimTraVI.Caption = gcPEN_SIMBOLO 'APRI20191022 SUGERENCIA CALIDAD
        lblMonTraVI.BackColor = &HC0FFFF
    Else
        lblSimTraVI.Caption = "$"
        lblMonTraVI.BackColor = &HC0FFC0
    End If
End Sub
Private Sub cboTransferMonedaVI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl Me.cmdTranferVI
    End If
End Sub
Private Sub cmdTranferVI_Click()
    Dim lsGlosa As String
    Dim lsDoc As String
    Dim lsInstit As String
    Dim oform As frmCapRegVouDepBus
    Dim lnTipMot As Integer
    Dim lsDetalle As String

    On Error GoTo ErrTransfer
    If cboTransferMonedaVI.Text = "" Then
        MsgBox "Debe escoger la moneda de la transferencia.", vbInformation, "Aviso"
        If cboTransferMonedaVI.Visible And cboTransferMonedaVI.Enabled Then EnfocaControl cboTransferMonedaVI
        Exit Sub
    End If

    lnTipMot = 16
    
    fnMovNroRVD = 0
    Set oform = New frmCapRegVouDepBus
    SetDatosTransferenciaVI "", "", "", 0, -1, ""
    oform.iniciarFormulario Trim(Right(cboTransferMonedaVI, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer, lsDetalle
    If fnMovNroRVD = 0 Then
        Exit Sub
    End If
    SetDatosTransferenciaVI lsGlosa, lsInstit, lsDoc, lnTransferSaldo, lnMovNroTransfer, lsDetalle
    Exit Sub
ErrTransfer:
    MsgBox "Ha sucedido un error al cargar los datos de la Transferencia", vbCritical, "Aviso"
End Sub
Private Sub SetDatosTransferenciaVI(ByVal psGlosa As String, ByVal psInstit As String, ByVal psDoc As String, ByVal pnTransferSaldo As Currency, ByVal pnMovNroTransfer As Long, ByVal psDetalle As String)
    Dim oPersona As New DCOMPersona
    Dim rsPersona As New ADODB.Recordset
    Dim row As Integer
    
    txtTransferGlosaVI.Text = psGlosa
    lbltransferBcoVI.Caption = psInstit
    lblTrasferNDVI.Caption = psDoc
    

    If pnMovNroTransfer <> -1 Then
        EnfocaControl txtTransferGlosaVI
    End If

    txtTransferGlosaVI.Locked = True
    lblMonTraVI = Format(pnTransferSaldo, "#,##0.00")
End Sub
'**********************************************************************************************
'SV*********************************************************************************************
Private Sub cboTransferMonedaSV_Click()
    If Right(cboTransferMonedaSV, 3) = Moneda.gMonedaNacional Then
        lblSimTraSV.Caption = gcPEN_SIMBOLO 'APRI20191022 SUGERENCIA CALIDAD
        lblMonTraSV.BackColor = &HC0FFFF
    Else
        lblSimTraSV.Caption = "$"
        lblMonTraSV.BackColor = &HC0FFC0
    End If
End Sub
Private Sub cboTransferMonedaSV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl Me.cmdTranferSV
    End If
End Sub
Private Sub cmdTranferSV_Click()
    Dim lsGlosa As String
    Dim lsDoc As String
    Dim lsInstit As String
    Dim oform As frmCapRegVouDepBus
    Dim lnTipMot As Integer
    Dim lsDetalle As String

    On Error GoTo ErrTransfer
    If cboTransferMonedaSV.Text = "" Then
        MsgBox "Debe escoger la moneda de la transferencia.", vbInformation, "Aviso"
        If cboTransferMonedaSV.Visible And cboTransferMonedaSV.Enabled Then EnfocaControl cboTransferMonedaSV
        Exit Sub
    End If

    lnTipMot = 16
    
    fnMovNroRVD = 0
    Set oform = New frmCapRegVouDepBus
    SetDatosTransferenciaSV "", "", "", 0, -1, ""
    oform.iniciarFormulario Trim(Right(cboTransferMonedaSV, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer, lsDetalle
    If fnMovNroRVD = 0 Then
        Exit Sub
    End If
    SetDatosTransferenciaSV lsGlosa, lsInstit, lsDoc, lnTransferSaldo, lnMovNroTransfer, lsDetalle
    Exit Sub
ErrTransfer:
    MsgBox "Ha sucedido un error al cargar los datos de la Transferencia", vbCritical, "Aviso"
End Sub
Private Sub SetDatosTransferenciaSV(ByVal psGlosa As String, ByVal psInstit As String, ByVal psDoc As String, ByVal pnTransferSaldo As Currency, ByVal pnMovNroTransfer As Long, ByVal psDetalle As String)
    Dim oPersona As New DCOMPersona
    Dim rsPersona As New ADODB.Recordset
    Dim row As Integer
    
    txtTransferGlosaSV.Text = psGlosa
    lbltransferBcoSV.Caption = psInstit
    lblTrasferNDSV.Caption = psDoc
    

    If pnMovNroTransfer <> -1 Then
        EnfocaControl txtTransferGlosaSV
    End If

    txtTransferGlosaSV.Locked = True
    lblMonTraSV = Format(pnTransferSaldo, "#,##0.00")
End Sub
'**********************************************************************************************
'VP*********************************************************************************************
Private Sub cboTransferMonedaVP_Click()
    If Right(cboTransferMonedaVP, 3) = Moneda.gMonedaNacional Then
        lblSimTraVP.Caption = gcPEN_SIMBOLO 'APRI20191022 SUGERENCIA CALIDAD
        lblMonTraVP.BackColor = &HC0FFFF
    Else
        lblSimTraVP.Caption = "$"
        lblMonTraVP.BackColor = &HC0FFC0
    End If
End Sub
Private Sub cboTransferMonedaVP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl Me.cmdTranferVP
    End If
End Sub
Private Sub cmdTranferVP_Click()
    Dim lsGlosa As String
    Dim lsDoc As String
    Dim lsInstit As String
    Dim oform As frmCapRegVouDepBus
    Dim lnTipMot As Integer
    Dim lsDetalle As String

    On Error GoTo ErrTransfer
    If cboTransferMonedaVP.Text = "" Then
        MsgBox "Debe escoger la moneda de la transferencia.", vbInformation, "Aviso"
        If cboTransferMonedaVP.Visible And cboTransferMonedaVP.Enabled Then EnfocaControl cboTransferMonedaVP
        Exit Sub
    End If

    lnTipMot = 16
    
    fnMovNroRVD = 0
    Set oform = New frmCapRegVouDepBus
    SetDatosTransferenciaVP "", "", "", 0, -1, ""
    oform.iniciarFormulario Trim(Right(cboTransferMonedaVP, 3)), lnTipMot, lsGlosa, lsInstit, lsDoc, lnTransferSaldo, fsPersCodTransfer, fnMovNroRVD, lnMovNroTransfer, lsDetalle
    If fnMovNroRVD = 0 Then
        Exit Sub
    End If
    SetDatosTransferenciaVP lsGlosa, lsInstit, lsDoc, lnTransferSaldo, lnMovNroTransfer, lsDetalle
    Exit Sub
ErrTransfer:
    MsgBox "Ha sucedido un error al cargar los datos de la Transferencia", vbCritical, "Aviso"
End Sub
Private Sub SetDatosTransferenciaVP(ByVal psGlosa As String, ByVal psInstit As String, ByVal psDoc As String, ByVal pnTransferSaldo As Currency, ByVal pnMovNroTransfer As Long, ByVal psDetalle As String)
    Dim oPersona As New DCOMPersona
    Dim rsPersona As New ADODB.Recordset
    Dim row As Integer
    
    txtTransferGlosaVP.Text = psGlosa
    lbltransferBcoVP.Caption = psInstit
    lblTrasferNDVP.Caption = psDoc
    

    If pnMovNroTransfer <> -1 Then
        EnfocaControl txtTransferGlosaVP
    End If

    txtTransferGlosaVP.Locked = True
    lblMonTraVP = Format(pnTransferSaldo, "#,##0.00")
End Sub
'**********************************************************************************************

Private Sub IniciaCombo(ByRef cboConst As ComboBox, nCapConst As ConstanteCabecera)
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsConst As New ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(nCapConst)
    Set clsGen = Nothing
    Do While Not rsConst.EOF
        cboConst.AddItem rsConst("cDescripcion") & space(100) & rsConst("nConsValor")
        rsConst.MoveNext
    Loop
    cboConst.ListIndex = 0
End Sub
'*****************************************************************************************************


Private Sub CmbForPagVC_Click()
    EstadoFormaPagoVC IIf(CmbForPagVC.ListIndex = -1, -1, CInt(Trim(Right(IIf(CmbForPagVC.Text = "", "-1", CmbForPagVC.Text), 10))))
    If CmbForPagVC.ListIndex <> -1 Then
        If CInt(Trim(Right(CmbForPagVC.Text, 10))) = gCVTipoPagoEfectivo Then
     
            Dim oformVou As New frmCapRegVouDepBus
            Dim lnTipMot As Integer
            Dim sGlosa As String
            Dim sIF As String
            Dim sVaucher As String
            Dim sPersCod As String
            Dim sNombre As String
            Dim sDireccion As String
            Dim sDocumento As String
            lnTipMot = 15
            cmdGuardar.Enabled = True
            Me.fraTranfereciaVC.Enabled = False
        ElseIf CInt(Trim(Right(CmbForPagVC.Text, 10))) = gCVTipoPagoCargoCta Then
            Dim sCuenta As String
            Dim sTempOpeCod As String
            Me.fraTranfereciaVC.Enabled = False
            sTempOpeCod = "300120"
           
            sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(gAhoCargoServicioRecaudo), sNumTarj, 232, False)
            If Val(Mid(sCuenta, 6, 3)) <> "232" And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            If Len(sCuenta) = 0 Then txtCuentaCargoVC.SetFocusAge: Exit Sub
            txtCuentaCargoVC.Age = Mid(sCuenta, 4, 2)
            txtCuentaCargoVC.Prod = Mid(sCuenta, 6, 3)
            txtCuentaCargoVC.Cuenta = Mid(sCuenta, 9, 18)
            Call txtCuentaCargoVC_KeyPress(13)
        ElseIf CInt(Trim(Right(CmbForPagVC.Text, 10))) = gCVTipoPagoVoucher Then
            cboTransferMonedaVC.Enabled = False
            Me.fraTranfereciaVC.Enabled = True
            cboTransferMonedaVC.ListIndex = IndiceListaCombo(cboTransferMonedaVC, 1)
        End If
    End If
End Sub

Private Sub CargaControlesVC()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R = oCons.RecuperaConstantes(gCVTipoPagoBase, , , 3)
  
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, CmbForPagVC)
    
    IniciaCombo cboTransferMonedaVC, gMoneda
    cboTransferMonedaVC.ListIndex = IndiceListaCombo(cboTransferMonedaVC, 1)
    Exit Sub
ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Function ValidaFormaPagoVC() As Boolean
Dim clsCap As New COMNCaptaGenerales.NCOMCaptaMovimiento
    ValidaFormaPagoVC = True
    If CmbForPagVC.ListIndex = -1 Then
        MsgBox "No se ha seleccionado la forma de pago. Verifique.", vbInformation, "¡Aviso!"
        EnfocaControl CmbForPagVC
        ValidaFormaPagoVC = False
    End If
    
    If CInt(Trim(Right(CmbForPagVC.Text, 10))) = gCVTipoPagoCargoCta And Len(txtCuentaCargoVC.NroCuenta) <> 18 Then
        MsgBox "Favor de ingresar la cuenta de ahorros a debitar", vbInformation, "¡Aviso!"
        EnfocaControl CmbForPagVC
        ValidaFormaPagoVC = False
    End If
        
    If CInt(Trim(Right(CmbForPagVC.Text, 10))) = gCVTipoPagoCargoCta Then
        If Not clsCap.ValidaSaldoCuenta(txtCuentaCargoVC.NroCuenta, CDbl(txtTotalVC.Text)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
            ValidaFormaPagoVC = False
        End If
    End If
    
    If CInt(Trim(Right(CmbForPagVC.Text, 10))) = gCVTipoPagoVoucher Then
        If lblTrasferNDVC.Caption = "" Then
            MsgBox "Debe ingresar un numero de transferencia.", vbInformation, "Aviso"
            EnfocaControl cmdTranferVC
            ValidaFormaPagoVC = False
        End If

        If CDbl(txtTotalVC.Text) <> CDbl(IIf(Trim(lblMonTraVC.Caption) = "", 0, lblMonTraVC.Caption)) And ValidaFormaPagoVC = True Then
            MsgBox "El monto de la operación no es igual al monto de la tranferencia.", vbInformation, "¡Aviso!"
            ValidaFormaPagoVC = False
        End If
     End If
    lnFormaPago = CInt(Trim(Right(CmbForPagVC.Text, 10)))
    lsCtaCargo = txtCuentaCargoVC.NroCuenta
                            
End Function



Private Sub txtCuentaCargoVC_KeyPress(KeyAscii As Integer)
    Dim sMsg As String
    If KeyAscii = 13 Then sMsg = ValidaCuentaACargo(txtCuentaCargoVC.NroCuenta)
    If Len(sMsg) > 0 Then
        MsgBox sMsg, vbInformation, "¡Aviso!"
        txtCuentaCargoVC.SetFocus
        Exit Sub
    End If
    Dim sMoneda As String
    pnMoneda = 0

    If Len(txtCuentaCargoVC.NroCuenta) = 18 Then
        If CInt(Mid(txtCuentaCargoVC.NroCuenta, 9, 1)) <> gMonedaNacional Then
            MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la operación de recaudo.", vbOKOnly + vbInformation, App.Title
             Exit Sub
        End If
    End If
    pnMoneda = gMonedaNacional
    ObtieneDatosCuenta txtCuentaCargoVC.NroCuenta
End Sub
Private Sub AsignaValorITFVC()
    pnITFCargoCta = 0#
    If CInt(Trim(Right(CmbForPagVC.Text, 10))) = gCVTipoPagoCargoCta Then
        If gITF.gbITFAplica Then
            pnITFCargoCta = Format(gITF.fgITFCalculaImpuesto(pnMontoPagarCargo), "#0.00")
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(pnITFCargoCta))
            If nRedondeoITF > 0 Then
                  pnITFCargoCta = Format(CCur(pnITFCargoCta) - nRedondeoITF, "#,##0.00")
            End If
         End If
     End If
End Sub

'/***FIN VC*****************************************************************************/
'****************************SV*************************************/
Private Sub EstadoFormaPagoSV(ByVal nFormaPago As Integer)
    LblNumDocSV.Caption = ""
    txtCuentaCargoSV.NroCuenta = ""
    Select Case nFormaPago
        Case -1
            LblNumDocSV.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargoSV.Visible = False
            cmdGuardar.Enabled = True
            fraTranfereciaSV.Enabled = False
            Call IniciarVouher
        Case gCVTipoPagoEfectivo
            txtCuentaCargoSV.Visible = False
            LblNumDocSV.Visible = False
            lblNroDocumento.Visible = False
            cmdGuardar.Enabled = True
            fraTranfereciaSV.Enabled = False
            Call IniciarVouher
        Case gCVTipoPagoCargoCta
            LblNumDocSV.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargoSV.Visible = True
            txtCuentaCargoSV.Enabled = True
            txtCuentaCargoSV.CMAC = gsCodCMAC
            txtCuentaCargoSV.Prod = Trim(Str(gCapAhorros))
            cmdGuardar.Enabled = False
            fraTranfereciaSV.Enabled = False
            Call IniciarVouher
        Case gCVTipoPagoVoucher
            txtCuentaCargoSV.Visible = False
            LblNumDocSV.Visible = False
            lblNroDocumento.Visible = False
            cmdGuardar.Enabled = True
            fraTranfereciaSV.Enabled = True
    End Select
End Sub
Private Sub CmbForPagSV_Click()
    EstadoFormaPagoSV IIf(CmbForPagSV.ListIndex = -1, -1, CInt(Trim(Right(IIf(CmbForPagSV.Text = "", "-1", CmbForPagSV.Text), 10))))
    If CmbForPagSV.ListIndex <> -1 Then
        If CInt(Trim(Right(CmbForPagSV.Text, 10))) = gCVTipoPagoEfectivo Then
     
            Dim oformVou As New frmCapRegVouDepBus
            Dim lnTipMot As Integer
            Dim sGlosa As String
            Dim sIF As String
            Dim sVaucher As String
            Dim sPersCod As String
            Dim sNombre As String
            Dim sDireccion As String
            Dim sDocumento As String
            lnTipMot = 15
            cmdGuardar.Enabled = True
            Me.fraTranfereciaSV.Enabled = False
        ElseIf CInt(Trim(Right(CmbForPagSV.Text, 10))) = gCVTipoPagoCargoCta Then
            Dim sCuenta As String
            Dim sTempOpeCod As String
            Me.fraTranfereciaSV.Enabled = False
            sTempOpeCod = "300120"
            
            sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(gAhoCargoServicioRecaudo), sNumTarj, 232, False)
            If Val(Mid(sCuenta, 6, 3)) <> "232" And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            If Len(sCuenta) = 0 Then txtCuentaCargoSV.SetFocusAge: Exit Sub
            txtCuentaCargoSV.Age = Mid(sCuenta, 4, 2)
            txtCuentaCargoSV.Prod = Mid(sCuenta, 6, 3)
            txtCuentaCargoSV.Cuenta = Mid(sCuenta, 9, 18)
            Call txtCuentaCargoSV_KeyPress(13)
        ElseIf CInt(Trim(Right(CmbForPagSV.Text, 10))) = gCVTipoPagoVoucher Then
            cboTransferMonedaSV.Enabled = False
            Me.fraTranfereciaSV.Enabled = True
            cboTransferMonedaSV.ListIndex = IndiceListaCombo(cboTransferMonedaSV, 1)
        End If
    End If
End Sub

Private Sub CargaControlesSV()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R = oCons.RecuperaConstantes(gCVTipoPagoBase, , , 3)
  
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, CmbForPagSV)
    
    IniciaCombo cboTransferMonedaSV, gMoneda
    cboTransferMonedaSV.ListIndex = IndiceListaCombo(cboTransferMonedaSV, 1)
    Exit Sub
ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Function ValidaFormaPagoSV() As Boolean
Dim clsCap As New COMNCaptaGenerales.NCOMCaptaMovimiento
    ValidaFormaPagoSV = True
    If CmbForPagSV.ListIndex = -1 Then
        MsgBox "No se ha seleccionado la forma de pago. Verifique.", vbInformation, "¡Aviso!"
        EnfocaControl CmbForPagSV
        ValidaFormaPagoSV = False
    End If
    
    If CInt(Trim(Right(CmbForPagSV.Text, 10))) = gCVTipoPagoCargoCta And Len(txtCuentaCargoSV.NroCuenta) <> 18 Then
        MsgBox "Favor de ingresar la cuenta de ahorros a debitar", vbInformation, "¡Aviso!"
        EnfocaControl CmbForPagSV
        ValidaFormaPagoSV = False
    End If
        
    If CInt(Trim(Right(CmbForPagSV.Text, 10))) = gCVTipoPagoCargoCta Then
        If Not clsCap.ValidaSaldoCuenta(txtCuentaCargoSV.NroCuenta, CDbl(txtTotalSV.Text)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
            ValidaFormaPagoSV = False
        End If
    End If
    If CInt(Trim(Right(CmbForPagSV.Text, 10))) = gCVTipoPagoVoucher Then
        If lblTrasferNDSV.Caption = "" Then
            MsgBox "Debe ingresar un numero de transferencia.", vbInformation, "Aviso"
            EnfocaControl cmdTranferSV
            ValidaFormaPagoSV = False
        End If
    
        If CDbl(txtTotalSV.Text) <> CDbl(IIf(Trim(lblMonTraSV.Caption) = "", 0, lblMonTraSV.Caption)) And ValidaFormaPagoSV = True Then
            MsgBox "El monto de la operación no es igual al monto de la tranferencia.", vbInformation, "¡Aviso!"
            ValidaFormaPagoSV = False
        End If
    End If
    lnFormaPago = CInt(Trim(Right(CmbForPagSV.Text, 10)))
    lsCtaCargo = txtCuentaCargoSV.NroCuenta
                            
End Function
Private Sub txtCuentaCargoSV_KeyPress(KeyAscii As Integer)
    Dim sMsg As String
    If KeyAscii = 13 Then sMsg = ValidaCuentaACargo(txtCuentaCargoSV.NroCuenta)
    If Len(sMsg) > 0 Then
        MsgBox sMsg, vbInformation, "¡Aviso!"
        txtCuentaCargoSV.SetFocus
        Exit Sub
    End If
    Dim sMoneda As String
    pnMoneda = 0

    If Len(txtCuentaCargoSV.NroCuenta) = 18 Then
        If CInt(Mid(txtCuentaCargoSV.NroCuenta, 9, 1)) <> gMonedaNacional Then
            MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la operación de recaudo.", vbOKOnly + vbInformation, App.Title
             Exit Sub
        End If
    End If
    pnMoneda = gMonedaNacional
    ObtieneDatosCuenta txtCuentaCargoSV.NroCuenta
End Sub
Private Sub AsignaValorITFSV()
    pnITFCargoCta = 0#
    If CInt(Trim(Right(CmbForPagSV.Text, 10))) = gCVTipoPagoCargoCta Then
        If gITF.gbITFAplica Then
            pnITFCargoCta = Format(gITF.fgITFCalculaImpuesto(pnMontoPagarCargo), "#0.00")
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(pnITFCargoCta))
            If nRedondeoITF > 0 Then
                  pnITFCargoCta = Format(CCur(pnITFCargoCta) - nRedondeoITF, "#,##0.00")
            End If
         End If
     End If
End Sub

'/***FIN SV*****************************************************************************/
'****************************VI*************************************/
Private Sub EstadoFormaPagoVI(ByVal nFormaPago As Integer)
    LblNumDocVI.Caption = ""
    txtCuentaCargoVI.NroCuenta = ""
    Select Case nFormaPago
        Case -1
            LblNumDocVI.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargoVI.Visible = False
            cmdGuardar.Enabled = True
            fraTranfereciaVI.Enabled = False
            Call IniciarVouher
        Case gCVTipoPagoEfectivo
            txtCuentaCargoVI.Visible = False
            LblNumDocVI.Visible = False
            lblNroDocumento.Visible = False
            lblNroDocumento.Visible = False
            cmdGuardar.Enabled = True
            fraTranfereciaVI.Enabled = False
            Call IniciarVouher
        Case gCVTipoPagoCargoCta
            LblNumDocVI.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargoVI.Visible = True
            txtCuentaCargoVI.Enabled = True
            txtCuentaCargoVI.CMAC = gsCodCMAC
            txtCuentaCargoVI.Prod = Trim(Str(gCapAhorros))
            cmdGuardar.Enabled = False
            fraTranfereciaVI.Enabled = False
            Call IniciarVouher
        Case gCVTipoPagoVoucher
            txtCuentaCargoVI.Visible = False
            LblNumDocVI.Visible = False
            lblNroDocumento.Visible = False
            cmdGuardar.Enabled = True
            fraTranfereciaVI.Enabled = True
    End Select
End Sub
Private Sub CmbForPagVI_Click()
    EstadoFormaPagoVI IIf(CmbForPagVI.ListIndex = -1, -1, CInt(Trim(Right(IIf(CmbForPagVI.Text = "", "-1", CmbForPagVI.Text), 10))))
    If CmbForPagVI.ListIndex <> -1 Then
        If CInt(Trim(Right(CmbForPagVI.Text, 10))) = gCVTipoPagoEfectivo Then
     
            Dim oformVou As New frmCapRegVouDepBus
            Dim lnTipMot As Integer
            Dim sGlosa As String
            Dim sIF As String
            Dim sVaucher As String
            Dim sPersCod As String
            Dim sNombre As String
            Dim sDireccion As String
            Dim sDocumento As String
            lnTipMot = 15
            cmdGuardar.Enabled = True
            Me.fraTranfereciaVI.Enabled = False
        ElseIf CInt(Trim(Right(CmbForPagVI.Text, 10))) = gCVTipoPagoCargoCta Then
            Dim sCuenta As String
            Dim sTempOpeCod As String
            Me.fraTranfereciaVI.Enabled = False
            sTempOpeCod = "300120"
            
            sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(gAhoCargoServicioRecaudo), sNumTarj, 232, False)
            If Val(Mid(sCuenta, 6, 3)) <> "232" And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            If Len(sCuenta) = 0 Then txtCuentaCargoVI.SetFocusAge: Exit Sub
            txtCuentaCargoVI.Age = Mid(sCuenta, 4, 2)
            txtCuentaCargoVI.Prod = Mid(sCuenta, 6, 3)
            txtCuentaCargoVI.Cuenta = Mid(sCuenta, 9, 18)
            Call txtCuentaCargoVI_KeyPress(13)
        ElseIf CInt(Trim(Right(CmbForPagVI.Text, 10))) = gCVTipoPagoVoucher Then
            cboTransferMonedaVI.Enabled = False
            Me.fraTranfereciaVI.Enabled = True
            cboTransferMonedaVI.ListIndex = IndiceListaCombo(cboTransferMonedaVI, 1)
        End If
    End If
End Sub

Private Sub CargaControlesVI()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R = oCons.RecuperaConstantes(gCVTipoPagoBase, , , 3)
  
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, CmbForPagVI)
    
    IniciaCombo cboTransferMonedaVI, gMoneda
    cboTransferMonedaVI.ListIndex = IndiceListaCombo(cboTransferMonedaVI, 1)
    Exit Sub
ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Function ValidaFormaPagoVI() As Boolean
Dim clsCap As New COMNCaptaGenerales.NCOMCaptaMovimiento
    ValidaFormaPagoVI = True
    If CmbForPagVI.ListIndex = -1 Then
        MsgBox "No se ha seleccionado la forma de pago. Verifique.", vbInformation, "¡Aviso!"
        EnfocaControl CmbForPagVI
        ValidaFormaPagoVI = False
    End If
    
    If CInt(Trim(Right(CmbForPagVI.Text, 10))) = gCVTipoPagoCargoCta And Len(txtCuentaCargoVI.NroCuenta) <> 18 Then
        MsgBox "Favor de ingresar la cuenta de ahorros a debitar", vbInformation, "¡Aviso!"
        EnfocaControl CmbForPagVI
        ValidaFormaPagoVI = False
    End If
        
    If CInt(Trim(Right(CmbForPagVI.Text, 10))) = gCVTipoPagoCargoCta Then
        If Not clsCap.ValidaSaldoCuenta(txtCuentaCargoVI.NroCuenta, CDbl(txtTotalVI.Text)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
            ValidaFormaPagoVI = False
        End If
    End If
    If CInt(Trim(Right(CmbForPagVI.Text, 10))) = gCVTipoPagoVoucher Then
        If lblTrasferNDVI.Caption = "" Then
            MsgBox "Debe ingresar un numero de transferencia.", vbInformation, "Aviso"
            EnfocaControl cmdTranferVI
            ValidaFormaPagoVI = False
        End If
    
        If CDbl(txtTotalVI.Text) <> CDbl(IIf(Trim(lblMonTraVI.Caption) = "", 0, lblMonTraVI.Caption)) And ValidaFormaPagoVI = True Then
            MsgBox "El monto de la operación no es igual al monto de la tranferencia.", vbInformation, "¡Aviso!"
            ValidaFormaPagoVI = False
        End If
    End If
    lnFormaPago = CInt(Trim(Right(CmbForPagVI.Text, 10)))
    lsCtaCargo = txtCuentaCargoVI.NroCuenta
                            
End Function
Private Sub txtCuentaCargoVI_KeyPress(KeyAscii As Integer)
    Dim sMsg As String
    If KeyAscii = 13 Then sMsg = ValidaCuentaACargo(txtCuentaCargoVI.NroCuenta)
    If Len(sMsg) > 0 Then
        MsgBox sMsg, vbInformation, "¡Aviso!"
        txtCuentaCargoVI.SetFocus
        Exit Sub
    End If
    Dim sMoneda As String
    pnMoneda = 0

    If Len(txtCuentaCargoVI.NroCuenta) = 18 Then
        If CInt(Mid(txtCuentaCargoVI.NroCuenta, 9, 1)) <> gMonedaNacional Then
            MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la operación de recaudo.", vbOKOnly + vbInformation, App.Title
             Exit Sub
        End If
    End If
    pnMoneda = gMonedaNacional
    ObtieneDatosCuenta txtCuentaCargoVI.NroCuenta
End Sub
Private Sub AsignaValorITFVI()
    pnITFCargoCta = 0#
    If CInt(Trim(Right(CmbForPagVI.Text, 10))) = gCVTipoPagoCargoCta Then
        If gITF.gbITFAplica Then
            pnITFCargoCta = Format(gITF.fgITFCalculaImpuesto(pnMontoPagarCargo), "#0.00")
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(pnITFCargoCta))
            If nRedondeoITF > 0 Then
                  pnITFCargoCta = Format(CCur(pnITFCargoCta) - nRedondeoITF, "#,##0.00")
            End If
         End If
     End If
End Sub
'/***FIN VI*****************************************************************************/
'****************************VP*************************************/
Private Sub EstadoFormaPagoVP(ByVal nFormaPago As Integer)
    LblNumDocVP.Caption = ""
    txtCuentaCargoVP.NroCuenta = ""
    Select Case nFormaPago
        Case -1
            LblNumDocVP.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargoVP.Visible = False
            cmdGuardar.Enabled = True
            fraTranfereciaVP.Enabled = False
            Call IniciarVouher
        Case gCVTipoPagoEfectivo
            txtCuentaCargoVP.Visible = False
            LblNumDocVP.Visible = False
            lblNroDocumento.Visible = False
            lblNroDocumento.Visible = False
            cmdGuardar.Enabled = True
            fraTranfereciaVP.Enabled = False
            Call IniciarVouher
        Case gCVTipoPagoCargoCta
            LblNumDocVP.Visible = False
            lblNroDocumento.Visible = False
            txtCuentaCargoVP.Visible = True
            txtCuentaCargoVP.Enabled = True
            txtCuentaCargoVP.CMAC = gsCodCMAC
            txtCuentaCargoVP.Prod = Trim(Str(gCapAhorros))
            cmdGuardar.Enabled = False
            fraTranfereciaVP.Enabled = False
            Call IniciarVouher
        Case gCVTipoPagoVoucher
            txtCuentaCargoVP.Visible = False
            LblNumDocVP.Visible = False
            lblNroDocumento.Visible = False
            cmdGuardar.Enabled = True
            fraTranfereciaVP.Enabled = True
    End Select
End Sub
Private Sub CmbForPagVP_Click()
    EstadoFormaPagoVP IIf(CmbForPagVP.ListIndex = -1, -1, CInt(Trim(Right(IIf(CmbForPagVP.Text = "", "-1", CmbForPagVP.Text), 10))))
    If CmbForPagVP.ListIndex <> -1 Then
        If CInt(Trim(Right(CmbForPagVI.Text, 10))) = gCVTipoPagoEfectivo Then
     
            Dim oformVou As New frmCapRegVouDepBus
            Dim lnTipMot As Integer
            Dim sGlosa As String
            Dim sIF As String
            Dim sVaucher As String
            Dim sPersCod As String
            Dim sNombre As String
            Dim sDireccion As String
            Dim sDocumento As String
            lnTipMot = 15
            cmdGuardar.Enabled = True
            Me.fraTranfereciaVP.Enabled = False
        ElseIf CInt(Trim(Right(CmbForPagVP.Text, 10))) = gCVTipoPagoCargoCta Then
            Dim sCuenta As String
            Dim sTempOpeCod As String
            Me.fraTranfereciaVP.Enabled = False
            sTempOpeCod = "300120"
            
            sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(gAhoCargoServicioRecaudo), sNumTarj, 232, False)
            If Val(Mid(sCuenta, 6, 3)) <> "232" And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            If Len(sCuenta) = 0 Then txtCuentaCargoVP.SetFocusAge: Exit Sub
            txtCuentaCargoVP.Age = Mid(sCuenta, 4, 2)
            txtCuentaCargoVP.Prod = Mid(sCuenta, 6, 3)
            txtCuentaCargoVP.Cuenta = Mid(sCuenta, 9, 18)
            Call txtCuentaCargoVP_KeyPress(13)
        ElseIf CInt(Trim(Right(CmbForPagVP.Text, 10))) = gCVTipoPagoVoucher Then
            cboTransferMonedaVP.Enabled = False
            Me.fraTranfereciaVP.Enabled = True
            cboTransferMonedaVP.ListIndex = IndiceListaCombo(cboTransferMonedaVP, 1)
        End If
    End If
End Sub

Private Sub CargaControlesVP()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R = oCons.RecuperaConstantes(gCVTipoPagoBase, , , 3)
  
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, CmbForPagVP)
    
    IniciaCombo cboTransferMonedaVP, gMoneda
    cboTransferMonedaVP.ListIndex = IndiceListaCombo(cboTransferMonedaVP, 1)
    Exit Sub
ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Function ValidaFormaPagoVP() As Boolean
Dim clsCap As New COMNCaptaGenerales.NCOMCaptaMovimiento
    ValidaFormaPagoVP = True
    If CmbForPagVP.ListIndex = -1 Then
        MsgBox "No se ha seleccionado la forma de pago. Verifique.", vbInformation, "¡Aviso!"
        EnfocaControl CmbForPagVP
        ValidaFormaPagoVP = False
    End If
    
    If CInt(Trim(Right(CmbForPagVP.Text, 10))) = gCVTipoPagoCargoCta And Len(txtCuentaCargoVP.NroCuenta) <> 18 Then
        MsgBox "Favor de ingresar la cuenta de ahorros a debitar", vbInformation, "¡Aviso!"
        EnfocaControl CmbForPagVP
        ValidaFormaPagoVP = False
    End If
        
    If CInt(Trim(Right(CmbForPagVP.Text, 10))) = gCVTipoPagoCargoCta Then
        If Not clsCap.ValidaSaldoCuenta(txtCuentaCargoVP.NroCuenta, CDbl(txtTotalVX.Text)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
            ValidaFormaPagoVP = False
        End If
    End If
    
    If CInt(Trim(Right(CmbForPagVP.Text, 10))) = gCVTipoPagoVoucher Then
        If lblTrasferNDVP.Caption = "" Then
            MsgBox "Debe ingresar un numero de transferencia.", vbInformation, "Aviso"
            EnfocaControl cmdTranferVP
            ValidaFormaPagoVP = False
        End If
    
        If CDbl(txtTotalVX.Text) <> CDbl(IIf(Trim(lblMonTraVP.Caption) = "", 0, lblMonTraVP.Caption)) And ValidaFormaPagoVP = True Then
            MsgBox "El monto de la operación no es igual al monto de la tranferencia.", vbInformation, "¡Aviso!"
            ValidaFormaPagoVP = False
        End If
        lnFormaPago = CInt(Trim(Right(CmbForPagVP.Text, 10)))
        lsCtaCargo = txtCuentaCargoVP.NroCuenta
    End If
End Function
Private Sub txtCuentaCargoVP_KeyPress(KeyAscii As Integer)
    Dim sMsg As String
    If KeyAscii = 13 Then sMsg = ValidaCuentaACargo(txtCuentaCargoVP.NroCuenta)
    If Len(sMsg) > 0 Then
        MsgBox sMsg, vbInformation, "¡Aviso!"
        txtCuentaCargoVP.SetFocus
        Exit Sub
    End If
    Dim sMoneda As String
    pnMoneda = 0

    If Len(txtCuentaCargoVP.NroCuenta) = 18 Then
        If CInt(Mid(txtCuentaCargoVP.NroCuenta, 9, 1)) <> gMonedaNacional Then
            MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la operación de recaudo.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
    End If
    pnMoneda = gMonedaNacional
    ObtieneDatosCuenta txtCuentaCargoVP.NroCuenta
End Sub
Private Sub AsignaValorITFVP()
    pnITFCargoCta = 0#
    If CInt(Trim(Right(CmbForPagVP.Text, 10))) = gCVTipoPagoCargoCta Then
        If gITF.gbITFAplica Then
            pnITFCargoCta = Format(gITF.fgITFCalculaImpuesto(pnMontoPagarCargo), "#0.00")
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(pnITFCargoCta))
            If nRedondeoITF > 0 Then
                  pnITFCargoCta = Format(CCur(pnITFCargoCta) - nRedondeoITF, "#,##0.00")
            End If
         End If
     End If
End Sub

'/***FIN VP*****************************************************************************/
Private Function ValidaCuentaACargo(ByVal psCuenta As String) As String
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sMsg As String
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    sMsg = clsCap.ValidaCuentaOperacion(psCuenta)
    ValidaCuentaACargo = sMsg
End Function
Private Sub ObtieneDatosCuenta(ByVal psCuenta As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsCta As ADODB.Recordset
    Dim rsV As ADODB.Recordset
    Dim lnTpoPrograma As Integer
    Dim lsTieneTarj As String
    Dim lbVistoVal As Boolean
    Dim lsOpeServicioRecaudoCargoCtaAhorro As String 'CTI06
    lsOpeServicioRecaudoCargoCtaAhorro = "" 'CTI06
    If pnMoneda = gMonedaNacional Then
        lsOpeServicioRecaudoCargoCtaAhorro = gAhoCargoServicioRecaudo
    Else
        lsOpeServicioRecaudoCargoCtaAhorro = ""
    End If

    If lsOpeServicioRecaudoCargoCtaAhorro = "" Then
        MsgBox "La moneda no coincide con la moneda de la operación. " & vbNewLine & "", vbInformation, "Aviso"
        Exit Sub
    End If
                        
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsV = New ADODB.Recordset
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(psCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
        If sNumTarj = "" Then
            If Mid(psCuenta, 6, 3) = "232" And lnTpoPrograma <> 1 Then
                Set rsV = clsMant.ValidaTarjetizacion(psCuenta, lsTieneTarj)
                If rsV.RecordCount > 0 Then
                    Dim tipoCta As Integer
                    tipoCta = rsCta("nPrdCtaTpo")
                    If tipoCta = 0 Or tipoCta = 2 Then
                        Dim rsCli As New ADODB.Recordset
                        Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
                        Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
                        Dim bExitoSol As Integer
                        Set rsCli = clsCli.GetPersonaCuenta(psCuenta, gCapRelPersTitular)
                        psPersCodTitularAhorroCargo = rsCli!cperscod ' CTI6
                        nRespuesta = oSolicitud.SolicitarVistoAtencionSinTarjeta(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(lsOpeServicioRecaudoCargoCtaAhorro))
                    
                        If nRespuesta = 1 Then '1:Tiene Visto de atencion sin tarjeta pendiente de autorizar
                             MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                             Exit Sub
                        End If
                        If nRespuesta = 2 Then '2:Tiene visto aceptado
                            MsgBox "La solicitud de atención sin tarjeta fue Aprobada, proceda con la atención", vbInformation, "Aviso"
                        End If
                        If nRespuesta = 3 Then '3:Tiene visto rechazado
                           If MsgBox("La solicitud de atención sin tarjeta fue RECHAZADA. ¿Desea realizar una nueva solicitud?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                Call oSolicitud.RegistrarVistoDeUsuario(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(lsOpeServicioRecaudoCargoCtaAhorro))
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        If nRespuesta = 4 Then '4:Se permite registrar la solicitud
                            Dim mensaje As String
                            If lsTieneTarj = "SI" Then
                                mensaje = "El Cliente posee tarjeta. Para continuar deberá registrar el Motivo de Autorización y comunicar al Coordinador o Jefe de Operaciones para su Aprobación. ¿Desea Continuar?"
                            Else
                                mensaje = "El Cliente NO posee tarjeta activa. Para continuar deberá registrar el Motivo de Autorización y comunicar al Coordinador o Jefe de Operaciones para su Aprobación. ¿Desea Continuar?"
                            End If
                        
                            If MsgBox(mensaje, vbInformation + vbYesNo, "Aviso") = vbYes Then
                        
                                bExitoSol = frmCapMotivoAutorizacion.Inicio(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(lsOpeServicioRecaudoCargoCtaAhorro))
                                If bExitoSol > 0 Then
                                    MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                                    Unload Me
                                End If
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        lbVistoVal = loVistoElectronico.Inicio(5, lsOpeServicioRecaudoCargoCtaAhorro)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema"
                            Exit Sub
                        End If
                    End If
                ElseIf lsTieneTarj = "NO" And rsV.RecordCount > 0 Then
                    If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? ", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg ers 065-2017
                        lbVistoVal = loVistoElectronico.Inicio(5, lsOpeServicioRecaudoCargoCtaAhorro)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema" 'ADD BY MARG ERS065-2017
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            End If
        Else
            If Mid(psCuenta, 6, 3) = "232" And lnTpoPrograma <> 1 Then
                Set rsV = clsMant.ValidaTarjetizacion(psCuenta, lsTieneTarj)
                If rsV.RecordCount > 0 Then
                    Dim tipoCta2 As Integer
                    tipoCta2 = rsCta("nPrdCtaTpo")
                    If tipoCta2 = 0 Or tipoCta2 = 2 Then
                        Dim rsCli2 As New ADODB.Recordset
                        Dim clsCli2 As New COMNCaptaGenerales.NCOMCaptaGenerales
                        Dim oSolicitud2 As New COMDCaptaGenerales.DCOMCaptaGenerales
                        Dim bExitoSol2 As Integer
                        Dim nRespuesta2 As Integer
                        Set rsCli2 = clsCli2.GetPersonaCuenta(psCuenta, gCapRelPersTitular)
                        psPersCodTitularAhorroCargo = rsCli2!cperscod ' CTI6
                    End If
                End If
            End If
        End If
        txtCuentaCargoVC.Enabled = False
        txtCuentaCargoVI.Enabled = False
        txtCuentaCargoSV.Enabled = False
        txtCuentaCargoVP.Enabled = False
        cmdGuardar.Enabled = True
        cmdGuardar.SetFocus
    End If
End Sub


'End CTI6 ERS0112020

'Para validar el DOI
Private Sub cboTipoDOISV_Click()
    If cboTipoDOISV.ListIndex = 0 Then
        txtDOISV.MaxLength = 8
        txtDOISV.Text = Empty
        txtDOISV.SetFocus
    ElseIf cboTipoDOISV.ListIndex = 1 Then
        txtDOISV.MaxLength = 11
        txtDOISV.Text = Empty
        If txtDOISV.Enabled Then txtDOISV.SetFocus
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 And bFocusGrid Then
        KeyCode = 10
    End If
End Sub
Private Sub cboTipoDOISV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDOISV.SetFocus
    End If
End Sub

'Para validar el DOI
Private Sub cboTipoDOIVX_Click()

    If cboTipoDOIVX.ListIndex = 0 Then
    
        txtDOIVX.MaxLength = 8
        txtDOIVX.Text = Empty
        txtDOIVX.SetFocus
    
    ElseIf cboTipoDOIVX.ListIndex = 1 Then
    
        txtDOIVX.MaxLength = 11
        txtDOIVX.Text = Empty
        txtDOIVX.SetFocus
    
    End If

End Sub

Private Sub cboTipoDOIVX_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtDOIVX.SetFocus
    End If

End Sub

Private Sub cmdAgregarSV_Click()

    grdConceptoPagarSV.AdicionaFila
    grdConceptoPagarSV.SetFocus
    grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.Rows - 1, 2) = IIf(nMoneda = 1, "SOLES", "DOLARES")
    grdConceptoPagarSV.row = grdConceptoPagarSV.Rows - 1
    grdConceptoPagarSV.Col = 1
    SendKeys "{F2}"
    
End Sub

Private Sub cmdAgregarSV_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdQuitarSV.SetFocus
    End If

End Sub

Private Sub cmdAgregarVX_Click()

    grdConceptoPagarVX.AdicionaFila
    grdConceptoPagarVX.SetFocus
    grdConceptoPagarVX.TextMatrix(grdConceptoPagarVX.Rows - 1, 2) = _
                                                                IIf(nMoneda = 1, "SOLES", "DOLARES")
    SendKeys "{ENTER}"
    
End Sub

Private Sub cmdBuscarPersonaVC_Click()
    
    Dim oBuscaUsuario As frmBuscarUsuarioConvenio
    Dim ClsServicioRecaudoWS As COMDCaptaServicios.DCOMSrvRecaudoWS 'CTI1 TI-ERS027-2019
    Set ClsServicioRecaudoWS = New COMDCaptaServicios.DCOMSrvRecaudoWS 'CTI1 TI-ERS027-2019
    Set oBuscaUsuario = New frmBuscarUsuarioConvenio
    Set rsUsuarioRecaudo = New Recordset
    Set rsUsuarioRecaudo = oBuscaUsuario.Inicio(Trim(txtCodigoBusConvenio.Text))
    
    If Not rsUsuarioRecaudo Is Nothing Then
        If Not (rsUsuarioRecaudo.EOF And rsUsuarioRecaudo.BOF) Then
            txtNombreClienteVC.Text = rsUsuarioRecaudo!cNomCliente
            txtDOIVC.Text = rsUsuarioRecaudo!cDOI
            txtCodigoIDVC.Text = rsUsuarioRecaudo!cCodCliente
        Else
            Exit Sub
        End If
    Dim nRow  As Integer
    LimpiaFlex grdConceptoPagarVC
    grdConceptoPagarVC_OnCellCheck 1, 5

        verificaWS = ClsServicioRecaudoWS.VerificarConvenioRecaudoWebService(Trim(txtCodigoBusConvenio.Text)) 'CTI1 TI-ERS027-2019
        If verificaWS = False Then 'CTI1 TI-ERS027-2019 begin
            Do While Not rsUsuarioRecaudo.EOF
                'Se valida el estado del cliente enviado por la empresa
                If rsUsuarioRecaudo!nEstado = Registrado Or rsUsuarioRecaudo!nEstado = Pagando Then

                    'ID-Servicio-Concepto-Fec Vec-Moneda-Importe-Mora-Pagar-a

                    grdConceptoPagarVC.AdicionaFila
                    nRow = grdConceptoPagarVC.Rows - 1

                    'Id
                    grdConceptoPagarVC.TextMatrix(nRow, 0) = rsUsuarioRecaudo!cId
                    'Servicio
                    grdConceptoPagarVC.TextMatrix(nRow, 1) = IIf(Len(Trim(rsUsuarioRecaudo!cServicio)) = 0, space(200) & ".", Trim(rsUsuarioRecaudo!cServicio))
                    'Concepto
                    grdConceptoPagarVC.TextMatrix(nRow, 2) = rsUsuarioRecaudo!cConcepto
                    'Fec Vec
                    grdConceptoPagarVC.TextMatrix(nRow, 3) = rsUsuarioRecaudo!dFechaVencimiento
                    'Moneda
                    grdConceptoPagarVC.TextMatrix(nRow, 4) = rsUsuarioRecaudo!cMoneda
                    'Importe
                    grdConceptoPagarVC.TextMatrix(nRow, 5) = rsUsuarioRecaudo!nImporte
                    grdConceptoPagarVC.TextMatrix(nRow, 5) = Format$(grdConceptoPagarVC.TextMatrix(nRow, 5), "#0.00")
                    'Mora
                    grdConceptoPagarVC.TextMatrix(nRow, 6) = rsUsuarioRecaudo!nMora
                    grdConceptoPagarVC.TextMatrix(nRow, 6) = Format$(grdConceptoPagarVC.TextMatrix(nRow, 6), "#0.00")
                    'copy ID
                    grdConceptoPagarVC.TextMatrix(nRow, 8) = rsUsuarioRecaudo!cId 'se copia valor para efectuar validaciones

                    'RIRO 20161102 ERS063-2016 ***
                    'nOrdenFecha
                    grdConceptoPagarVC.TextMatrix(nRow, 9) = rsUsuarioRecaudo!nOrdenFecha 'Muestra el orden de la fecha de vencimiento.
                    'END RIRO ********************

                End If
                rsUsuarioRecaudo.MoveNext
            Loop
            rsUsuarioRecaudo.MoveFirst
        Else 'CTI1 TI-ERS027-2019 begin
            Do While Not rsUsuarioRecaudo.EOF
                grdConceptoPagarVC.AdicionaFila
                nRow = grdConceptoPagarVC.Rows - 1
                'Id
                grdConceptoPagarVC.TextMatrix(nRow, 0) = rsUsuarioRecaudo!cCodigoComprobante 'YA
                'Servicio
                grdConceptoPagarVC.TextMatrix(nRow, 1) = IIf(Len(Trim(cServicioWS)) = 0, space(200) & ".", Trim(cServicioWS))
                'Concepto
                grdConceptoPagarVC.TextMatrix(nRow, 2) = rsUsuarioRecaudo!cConcepto 'YA
                'Fec Vec
                grdConceptoPagarVC.TextMatrix(nRow, 3) = rsUsuarioRecaudo!cFechaVencimiento 'YA
                'Moneda
                grdConceptoPagarVC.TextMatrix(nRow, 4) = cMonedaWS
                'Importe
                grdConceptoPagarVC.TextMatrix(nRow, 5) = rsUsuarioRecaudo!nMontoTotal 'YA
                grdConceptoPagarVC.TextMatrix(nRow, 5) = Format$(grdConceptoPagarVC.TextMatrix(nRow, 5), "#0.00")
                'Mora
                grdConceptoPagarVC.TextMatrix(nRow, 6) = rsUsuarioRecaudo!nMora
                grdConceptoPagarVC.TextMatrix(nRow, 6) = Format$(grdConceptoPagarVC.TextMatrix(nRow, 6), "#0.00")
                'copy ID
                grdConceptoPagarVC.TextMatrix(nRow, 8) = rsUsuarioRecaudo!cCodigoComprobante 'se copia valor para efectuar validaciones

                'RIRO 20161102 ERS063-2016 ***
                'nOrdenFecha
                'grdConceptoPagarVC.TextMatrix(nRow, 9) = rsUsuarioRecaudo!nOrdenFecha 'Muestra el orden de la fecha de vencimiento.

                rsUsuarioRecaudo.MoveNext
            Loop
            rsUsuarioRecaudo.MoveFirst
        End If 'CTI1 TI-ERS027-2019 begin
        CmbForPagVC.Enabled = True 'CTI6 ERS0112020
        CmbForPagVC.ListIndex = IndiceListaCombo(CmbForPagVC, 1) 'CTI6 ERS0112020
Else
    MsgBox "Usted no selecciono ningun usuario", vbExclamation, "Aviso"
    limpiaDetalle
End If

End Sub

Private Sub cmdBuscarPersonaVI_Click()
        
    Set rsUsuarioRecaudo = New Recordset
    Set rsUsuarioRecaudo = frmBuscarUsuarioConvenio.Inicio(Trim(txtCodigoBusConvenio.Text))
    
    If Not rsUsuarioRecaudo Is Nothing Then
    
        If Not (rsUsuarioRecaudo.EOF And rsUsuarioRecaudo.BOF) Then

            txtNombreClienteVI.Text = rsUsuarioRecaudo!cNomCliente
            txtDOIVI.Text = rsUsuarioRecaudo!cDOI
            txtCodigoIDVI.Text = rsUsuarioRecaudo!cCodCliente

        End If
        
        Dim nRow  As Integer
        LimpiaFlex grdConceptoPagarVI
        calculoSubTotalComisioVI
        Do While Not rsUsuarioRecaudo.EOF
        
            grdConceptoPagarVI.AdicionaFila
            
            nRow = grdConceptoPagarVI.Rows - 1
            
            grdConceptoPagarVI.TextMatrix(nRow, 0) = rsUsuarioRecaudo!cId
            grdConceptoPagarVI.TextMatrix(nRow, 1) = rsUsuarioRecaudo!cServicio
            grdConceptoPagarVI.TextMatrix(nRow, 2) = rsUsuarioRecaudo!cConcepto
            grdConceptoPagarVI.TextMatrix(nRow, 3) = rsUsuarioRecaudo!cMoneda
            grdConceptoPagarVI.TextMatrix(nRow, 5) = Format(rsUsuarioRecaudo!nDeudaActual, "#,##0.00")
            grdConceptoPagarVI.TextMatrix(nRow, 6) = Format(rsUsuarioRecaudo!nDeudaActual, "#,##0.00") 'RIRO20170623
            'grdConceptoPagarVI.TextMatrix(nRow, 5) = Format$(grdConceptoPagarVI.TextMatrix(nRow, 5), "#,##0.00")
            
            ReDim Preserve Importes(nRow)
            Importes(nRow) = CDbl(rsUsuarioRecaudo!nDeudaActual)
            
            rsUsuarioRecaudo.MoveNext
            
        Loop
        rsUsuarioRecaudo.MoveFirst
         
        CmbForPagVI.Enabled = True 'CTI6 ERS0112020
        CmbForPagVI.ListIndex = IndiceListaCombo(CmbForPagVI, 1) 'CTI6 ERS0112020
        
    Else
        MsgBox "Usted no selecciono ninguna Empresa", vbExclamation, "Aviso"
         limpiaDetalle
         
    End If

End Sub

Private Sub cmdBuscarConvenio_Click()

    Dim rsRecaudo As Recordset
    Set rsRecaudo = New Recordset
    Set rsRecaudo = frmBuscarConvenio.Inicio

    ' Limpia Detalle de Convevio
    limpiaDetalle
    Set rsUsuarioRecaudo = Nothing
    nMoneda = 1
    nComisionEmpresa = 0#
    strCuenta = ""
    
    If Not rsRecaudo Is Nothing Then
        If Not (rsRecaudo.EOF And rsRecaudo.BOF) Then
            
            Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
            Dim objValidar As COMNCaptaGenerales.NCOMCaptaMovimiento
            Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
            Set objValidar = New COMNCaptaGenerales.NCOMCaptaMovimiento
            
            'Valida el estado de la cuenta, solo permite el pago a cuentas activas
            If Not objValidar.ValidaEstadoCuenta(rsRecaudo!cCtaCod, False) Then
                MsgBox "La cuenta del convenio seleccionado NO Tiene un estado valido para la operacion", vbExclamation + vbDefaultButton1, "Aviso"
                limpiaDetalle
                limipiarCabecera
                Set ClsServicioRecaudo = Nothing
                Set objValidar = Nothing
                Exit Sub
                
            End If

            'CTI1 TI-ERS027-2019 begin
            Dim ClsServicioRecaudoWS As COMDCaptaServicios.DCOMSrvRecaudoWS
            Set ClsServicioRecaudoWS = New COMDCaptaServicios.DCOMSrvRecaudoWS
            verificaWS = ClsServicioRecaudoWS.VerificarConvenioRecaudoWebService(rsRecaudo!cCodConvenio)

            If verificaWS = False Then
                'Verifica si en el convenio seleccionado, hay clientes a pagar
                If Mid(rsRecaudo!cCodConvenio, 14, 2) <> "SV" Then
                    If Not ClsServicioRecaudo.getCantidadRegistrosConvenio(rsRecaudo!cCodConvenio) Then
                        If Mid(rsRecaudo!cCodConvenio, 14, 2) = "VP" Then
                            MsgBox "El convenio seleccionado no posee importes a pagar", vbExclamation, "Aviso"
                        Else
                            MsgBox "El convenio seleccionado no posee clientes a pagar", vbExclamation, "Aviso"
                        End If
                        limpiaDetalle
                        limipiarCabecera
                        Set ClsServicioRecaudo = Nothing
                        Set objValidar = Nothing
                        Exit Sub
                    End If
                End If
            Else 'CTI1 TI-ERS027-2019 end
                cMonedaWS = rsRecaudo!cMoneda
                cServicioWS = rsRecaudo!cServicio
            End If 'CTI1 TI-ERS027-2019 end
            
            txtCodigoBusConvenio.Text = rsRecaudo!cCodConvenio
            txtNombreConvenio.Text = rsRecaudo!cNombreConvenio
            txtCodigoEmpresa.Text = rsRecaudo!cperscod
            txtNombreEmpresa.Text = rsRecaudo!cPersNombre
            nMoneda = IIf(Mid(rsRecaudo!cCtaCod, 9, 1) = 1, 1, 2)
            strCuenta = rsRecaudo!cCtaCod
            
            lblDescripcionCodigoVC.Caption = rsRecaudo!cDescripcion
            lblDescripcionCodigoVI.Caption = rsRecaudo!cDescripcion
            lblDescripcionCodigoSV.Caption = rsRecaudo!cDescripcion
            lblDescripcionCodigoVP.Caption = rsRecaudo!cDescripcion
            
            vistaPestana (txtCodigoBusConvenio.Text)
            
            If Mid(txtCodigoBusConvenio.Text, 14, 2) = "VP" Then
                grdConceptoPagarVX.CargaCombo ClsServicioRecaudo.getListaConceptosCobrarXConvenioPV(Trim(txtCodigoBusConvenio.Text))
            End If
            
            txtCodigoBusConvenio.Locked = True
            cmdGuardar.Enabled = True
            cmdcancelar.Enabled = True
            Set ClsServicioRecaudo = Nothing
            Set objValidar = Nothing
            
        End If
        InicializarComboPago 'CTI6 ERS0112020
    Else
         MsgBox "No selecciono ningun convenio", vbExclamation, "Aviso"
         limipiarCabecera
    End If

End Sub

Private Sub InicializarComboPago()
 CmbForPagSV.Enabled = True 'CTI6 ERS0112020
 CmbForPagSV.ListIndex = IndiceListaCombo(CmbForPagSV, 1) 'CTI6 ERS0112020
 
 CmbForPagVC.Enabled = True 'CTI6 ERS0112020
 CmbForPagVC.ListIndex = IndiceListaCombo(CmbForPagVC, 1) 'CTI6 ERS0112020
 
 CmbForPagVI.Enabled = True 'CTI6 ERS0112020
 CmbForPagVI.ListIndex = IndiceListaCombo(CmbForPagVI, 1) 'CTI6 ERS0112020
 
 CmbForPagVP.Enabled = True 'CTI6 ERS0112020
 CmbForPagVP.ListIndex = IndiceListaCombo(CmbForPagVP, 1) 'CTI6 ERS0112020
End Sub

'Limpia la parte superior del formulario
Private Sub limipiarCabecera()

    txtCodigoBusConvenio.Text = ""
     txtNombreConvenio.Text = ""
     txtCodigoEmpresa.Text = ""
     txtNombreEmpresa.Text = ""
     strCuenta = ""
     stContenedorValidacion.Enabled = False
     txtCodigoBusConvenio.SetFocus
     txtCodigoBusConvenio.Locked = False
     cmdGuardar.Enabled = False
     cmdcancelar.Enabled = False
     
End Sub
     
'Limpia la parte central del formulario
Private Sub limpiaDetalle()
    
    Set rsUsuarioRecaudo = Nothing
    nMoneda = 1
    nComisionEmpresa = 0#
    
' ***** Validacion Completa *****
    'grdConceptoPagarVC.Clear
    grdConceptoPagarVC.Rows = 2
    LimpiaFlex grdConceptoPagarVC
    txtCodigoIDVC.Text = ""
    txtNombreClienteVC.Text = ""
    txtDOIVC.Text = ""
    txtSubTotalVC.Text = "0.00"
    txtComisionVC.Text = "0.00"
    txtTotalVC.Text = "0.00"
    
' ***** Validacion Incompleta *****
    grdConceptoPagarVI.Clear
    grdConceptoPagarVI.Rows = 2
    grdConceptoPagarVI.FormaCabecera
    txtCodigoIDVI.Text = ""
    txtNombreClienteVI.Text = ""
    txtDOIVI.Text = ""
    txtSubTotalVI.Text = "0.00"
    txtComisionVI.Text = "0.00"
    txtTotalVI.Text = "0.00"

' ***** Sin Validacion *****
    grdConceptoPagarSV.Clear
    grdConceptoPagarSV.Rows = 2
    grdConceptoPagarSV.FormaCabecera
    txtOtroCodigoSV.Text = ""
    txtNombreClienteSV.Text = ""
    txtDOISV.Text = ""
    cboTipoDOISV.ListIndex = 1
    txtSubTotalSV.Text = "0.00"
    txtComisionSV.Text = "0.00"
    txtTotalSV.Text = "0.00"
    
' ***** Validacion por Importe *****
    grdConceptoPagarVX.Clear
    grdConceptoPagarVX.Rows = 2
    grdConceptoPagarVX.FormaCabecera
    txtOtroCodigoVX.Text = ""
    txtNombreClienteVX.Text = ""
    txtDOIVX.Text = ""
    txtSubTotalVX.Text = "0.00"
    txtComisionVX.Text = "0.00"
    txtTotalVX.Text = "0.00"
    
    ' ***** Cargo a cuenta *****
   
    LblNumDocVC.Caption = "" 'CTI6 ERS0112020
    pnMoneda = 0 'CTI6 ERS0112020
    'VC
    CmbForPagVC.Enabled = False 'CTI6 ERS0112020
    CmbForPagVC.ListIndex = -1 'CTI6 ERS0112020
    txtCuentaCargoVC.NroCuenta = "" 'CTI6 ERS0112020
    
    CmbForPagSV.Enabled = False 'CTI6 ERS0112020
    CmbForPagSV.ListIndex = -1 'CTI6 ERS0112020
    txtCuentaCargoSV.NroCuenta = "" 'CTI6 ERS0112020
    
    CmbForPagVI.Enabled = False 'CTI6 ERS0112020
    CmbForPagVI.ListIndex = -1 'CTI6 ERS0112020
    txtCuentaCargoVI.NroCuenta = "" 'CTI6 ERS0112020
    
    CmbForPagVP.Enabled = False 'CTI6 ERS0112020
    CmbForPagVP.ListIndex = -1 'CTI6 ERS0112020
    txtCuentaCargoVP.NroCuenta = "" 'CTI6 ERS0112020
    
    IniciarVouher 'CTI7 OPEv2
    
    ReDim Preserve Importes(0)
    
End Sub
Private Sub IniciarVouher()
    Me.fraTranfereciaVC.Enabled = False 'CTI7 OPEv2
    cboTransferMonedaVC.ListIndex = -1 'CTI7 OPEv2
    lblTrasferNDVC.Caption = "" 'CTI7 OPEv2
    lbltransferBcoVC.Caption = "" 'CTI7 OPEv2
    txtTransferGlosaVC.Text = "" 'CTI7 OPEv2
    lblMonTraVC.Caption = "" 'CTI7 OPEv2
    
    Me.fraTranfereciaVI.Enabled = False 'CTI7 OPEv2
    cboTransferMonedaVI.ListIndex = -1 'CTI7 OPEv2
    lblTrasferNDVI.Caption = "" 'CTI7 OPEv2
    lbltransferBcoVI.Caption = "" 'CTI7 OPEv2
    txtTransferGlosaVI.Text = "" 'CTI7 OPEv2
    lblMonTraVI.Caption = "" 'CTI7 OPEv2
    
    Me.fraTranfereciaSV.Enabled = False 'CTI7 OPEv2
    cboTransferMonedaSV.ListIndex = -1 'CTI7 OPEv2
    lblTrasferNDSV.Caption = "" 'CTI7 OPEv2
    lbltransferBcoSV.Caption = "" 'CTI7 OPEv2
    txtTransferGlosaSV.Text = "" 'CTI7 OPEv2
    lblMonTraSV.Caption = "" 'CTI7 OPEv2
    
    Me.fraTranfereciaVP.Enabled = False 'CTI7 OPEv2
    cboTransferMonedaVP.ListIndex = -1 'CTI7 OPEv2
    lblTrasferNDVP.Caption = "" 'CTI7 OPEv2
    lbltransferBcoVP.Caption = "" 'CTI7 OPEv2
    txtTransferGlosaVP.Text = "" 'CTI7 OPEv2
    lblMonTraVP.Caption = "" 'CTI7 OPEv2
End Sub


Private Sub cmdCancelar_Click()

    If MsgBox("¿Está seguro de cancelar la operacion?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    
        limpiaDetalle
        limipiarCabecera
        
    End If
    
End Sub

Private Sub Cmdguardar_Click()
    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII
    
    Dim nI As Double
    Dim nCount As Double, nTmpValid As Integer
    Dim sMensaje As String
    Dim objValidar As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set objValidar = New COMNCaptaGenerales.NCOMCaptaMovimiento
    
    Dim MatDatosAho(14) As String 'CTI6 ERS0112020

    On Error GoTo Error

    If txtCodigoBusConvenio.Text = "" Or txtNombreConvenio.Text = "" Or _
        txtCodigoEmpresa.Text = "" Or txtNombreEmpresa.Text = "" Then
        
        MsgBox "No se Selecciono Ningun Convenio", vbExclamation, "Aviso"
        Exit Sub
                
    End If
    
    ' ===== Valida el estado de la cuenta
    If Not objValidar.ValidaEstadoCuenta(strCuenta, False) Then
        MsgBox "Cuenta NO Tiene un estado valido para la Operacion, consulte con el Asistente de Agencia.", vbExclamation, "aviso"
        Exit Sub
    End If
        
    If nConvenioSeleccionado = 0 Then
            MsgBox "No se selecciono ningun convenio", vbExclamation, "Aviso"
            Exit Sub
    ' Validando Convenio:  Sin Validacion
    ElseIf nConvenioSeleccionado = Convenio_SV Then
        If cboTipoDOISV.ListIndex = 0 Then
            If Len(Trim(txtDOISV.Text)) < 8 Then
                MsgBox "El tipo de documento seleccionado admite 8 caracteres", vbExclamation, "Aviso"
                Exit Sub
            End If
        ElseIf cboTipoDOISV.ListIndex = 1 Then
            If Len(Trim(txtDOISV.Text)) < 11 Then
                MsgBox "El tipo de documento seleccionado admite 11 caracteres", vbExclamation, "Aviso"
                Exit Sub
            End If
        End If
        sMensaje = Trim(validaSVVP)
        If Not sMensaje = "" Then
            MsgBox sMensaje, vbExclamation, "Aviso"
            Exit Sub
        End If
        'Validadndo Pago de Minims y Maximos
        
        
    ' Validando Convenio:  Validacion por Importes
    ElseIf nConvenioSeleccionado = Convenio_VP Then
            If cboTipoDOIVX.ListIndex = 0 Then
                If Len(Trim(txtDOIVX.Text)) < 8 Then
                    MsgBox "El tipo de documento seleccionado admite 8 caracteres", vbExclamation, "Aviso"
                    Exit Sub
                End If
            ElseIf cboTipoDOIVX.ListIndex = 1 Then
                If Len(Trim(txtDOIVX.Text)) < 11 Then
                    MsgBox "El tipo de documento seleccionado admite 11 caracteres", vbExclamation, "Aviso"
                    Exit Sub
                End If
            End If
            sMensaje = Trim(validaSVVP)
            If Not sMensaje = "" Then
                MsgBox sMensaje, vbExclamation, "Aviso"
                Exit Sub
            End If
    ' Validando Convenio:  Validacion Completa
    ElseIf nConvenioSeleccionado = Convenio_VC Or nConvenioSeleccionado = Convenio_VCM Then
            For nI = 1 To grdConceptoPagarVC.Rows - 1
                'If grdConceptoPagarVC.TextMatrix(nI, 5) = "." Then
                If grdConceptoPagarVC.TextMatrix(nI, 7) = "." Then
                    nCount = nCount + 1
                    If grdConceptoPagarVC.TextMatrix(nI, 0) <> grdConceptoPagarVC.TextMatrix(nI, 8) Then
                        nTmpValid = nTmpValid + 1
                    End If
                End If
            Next
            If nCount = 0 Then
                MsgBox "No se seleccionaron montos a pagar", vbExclamation, "Aviso"
                Exit Sub
            End If
            nCount = 0
            If nTmpValid > 0 Then
                MsgBox "Los valores del campo ID cambiaron, vuelva a cargar los valores", vbExclamation, "Aviso"
                Exit Sub
            End If
            nTmpValid = 0
            If rsUsuarioRecaudo Is Nothing Then
                MsgBox "No se seleccionó un usuario del convenio", vbExclamation, "Aviso"
                Exit Sub
            End If

            For nI = 1 To grdConceptoPagarVC.Rows - 1
                'If val(grdConceptoPagarVC.TextMatrix(nI, 4)) <= 0 Then
                If Val(grdConceptoPagarVC.TextMatrix(nI, 5)) <= 0 Then
                    nCount = nCount + 1
                End If
            Next
            If nCount > 0 Then
                MsgBox "No puede efectuar pagos que sean iguales a 0.00 ", vbExclamation, "Aviso"
                Exit Sub
            End If
    ' Validando Convenio:  Validacion Incompleta
    ElseIf nConvenioSeleccionado = Convenio_VI Then
            For nI = 1 To grdConceptoPagarVI.Rows - 1
                If grdConceptoPagarVI.TextMatrix(nI, 4) = "." Then
                    nCount = nCount + 1
                End If
            Next
            If rsUsuarioRecaudo Is Nothing Then
                MsgBox "No se seleccionó un usuario del convenio", vbExclamation, "Aviso"
                Exit Sub
            End If
            If nCount = 0 Then
                MsgBox "No se seleccionaron montos a pagar", vbExclamation, "Aviso"
                Exit Sub
            End If
            nCount = 0
            
            ' Aplicando Validacion de Importe Minimos y Maximos
            Dim rsImpoteMinMax As ADODB.Recordset
            Dim nMinimo, nMaximo As Double
                   
            For nI = 1 To grdConceptoPagarVI.Rows - 1
                If grdConceptoPagarVI.TextMatrix(nI, 4) = "." Then
                    If Val(grdConceptoPagarVI.TextMatrix(nI, 5)) <= 0 Then
                        nCount = nCount + 1
                    End If
                    
                    Set rsImpoteMinMax = ObtenerImporteMinMax(Trim(txtCodigoBusConvenio.Text), grdConceptoPagarVI.TextMatrix(nI, 0))
                    If Not rsImpoteMinMax Is Nothing Then
                        If Not rsImpoteMinMax.EOF And Not rsImpoteMinMax.BOF Then
                            nMinimo = rsImpoteMinMax!nPagoMin
                            nMaximo = rsImpoteMinMax!nPagoMax
                        Else
                            nMinimo = 0
                            nMaximo = 9999
                        End If
                    Else
                            nMinimo = 0
                            nMaximo = 9999
                    End If
                    
                    If CDbl(grdConceptoPagarVI.TextMatrix(nI, 5)) < nMinimo Or _
                       CDbl(grdConceptoPagarVI.TextMatrix(nI, 5)) > nMaximo Then
                       
                       MsgBox " Verificar importes. Límites por registro: Monto mínimo " & nMinimo & " - Monto máximo " & nMaximo _
                       , vbExclamation, "Aviso"
                                         
                       Set rsImpoteMinMax = Nothing
                       Exit Sub
                        
                    End If
                    
                End If
            Next
            If nCount > 0 Then
                MsgBox "No puede efectuar pagos que sean iguales a 0.00 ", vbExclamation, "Aviso"
                Exit Sub
            End If
    End If
    
    'RIRO 20161102 ********************
    If nConvenioSeleccionado = Convenio_VCM Then
        sMensaje = ValidarFechaPagos
        If Len(Trim(sMensaje)) > 0 Then
            MsgBox sMensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    'END RIRO *************************
    
    ' ===== Fin Validacion
    
    'Procediendo a grabar
    Dim pnForPago As Integer
    Dim pnMonedaTransferencia As Integer
    Dim pnMonTransferencia As Currency
    lnFormaPago = 0
    lsCtaCargo = ""
                         
    If MsgBox("¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        
        Dim sImpresion As String
        Dim sMovNro As String, sMovNro2 As String
        Dim oCont As COMNContabilidad.NCOMContFunciones
        Set oCont = New COMNContabilidad.NCOMContFunciones
        sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Sleep 1000
        sMovNro2 = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        If nConvenioSeleccionado = Convenio_VC Or nConvenioSeleccionado = Convenio_VCM Then
            pnForPago = CInt(Trim(Right(CmbForPagVC.Text, 10)))
            If Trim(Me.cboTransferMonedaVC.Text) = "" Then
                pnMonedaTransferencia = gMonedaNacional
            Else
                pnMonedaTransferencia = CInt(Right(Me.cboTransferMonedaVC.Text, 3))
            End If
            If Trim(lblMonTraVC) = "" Then
                pnMonTransferencia = 0
            Else
                pnMonTransferencia = CCur(lblMonTraVC)
            End If
            registrarVC sMovNro, sMovNro2, pnITFCargoCta, pnMontoPagarCargo, lnMovNroTransfer, pnMonedaTransferencia, fnMovNroRVD, CCur(pnMonTransferencia)
        ElseIf nConvenioSeleccionado = Convenio_VI Then
            pnForPago = CInt(Trim(Right(CmbForPagVI.Text, 10)))
            If Trim(Me.cboTransferMonedaVI.Text) = "" Then
                pnMonedaTransferencia = gMonedaNacional
            Else
                pnMonedaTransferencia = CInt(Right(Me.cboTransferMonedaVI.Text, 3))
            End If
            If Trim(lblMonTraVI) = "" Then
                pnMonTransferencia = 0
            Else
                pnMonTransferencia = CCur(lblMonTraVI)
            End If
            registrarVI sMovNro, sMovNro2, pnITFCargoCta, pnMontoPagarCargo, lnMovNroTransfer, pnMonedaTransferencia, fnMovNroRVD, CCur(pnMonTransferencia)
        ElseIf nConvenioSeleccionado = Convenio_SV Then
            pnForPago = CInt(Trim(Right(CmbForPagSV.Text, 10)))
            If Trim(Me.cboTransferMonedaSV.Text) = "" Then
                pnMonedaTransferencia = gMonedaNacional
            Else
                pnMonedaTransferencia = CInt(Right(Me.cboTransferMonedaSV.Text, 3))
            End If
            If Trim(lblMonTraSV) = "" Then
                pnMonTransferencia = 0
            Else
                pnMonTransferencia = CCur(lblMonTraSV)
            End If
            registrarSV sMovNro, sMovNro2, pnITFCargoCta, pnMontoPagarCargo, lnMovNroTransfer, pnMonedaTransferencia, fnMovNroRVD, CCur(pnMonTransferencia)
        ElseIf nConvenioSeleccionado = Convenio_VP Then
            pnForPago = CInt(Trim(Right(CmbForPagVP.Text, 10)))
            If Trim(Me.cboTransferMonedaVP.Text) = "" Then
                pnMonedaTransferencia = gMonedaNacional
            Else
                pnMonedaTransferencia = CInt(Right(Me.cboTransferMonedaVP.Text, 3))
            End If
            If Trim(lblMonTraVP) = "" Then
                pnMonTransferencia = 0
            Else
                pnMonTransferencia = CCur(lblMonTraVP)
            End If
            registrarVP sMovNro, sMovNro2, pnITFCargoCta, pnMontoPagarCargo, lnMovNroTransfer, pnMonedaTransferencia, fnMovNroRVD, CCur(pnMonTransferencia)
        End If
        
        'CTI4 ERS0112020
       ' If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
        If lnFormaPago = gColocTipoPagoCargoCta Then
            Dim oMovOperacion As COMDMov.DCOMMov
            Dim nMovNroOperacion As Long
            Dim rsCli As New ADODB.Recordset
            Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
            Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
            Set oMovOperacion = New COMDMov.DCOMMov
            nMovNroOperacion = oMovOperacion.GetnMovNro(sMovNro)

            loVistoElectronico.RegistraVistoElectronico nMovNroOperacion, , gsCodUser, nMovNroOperacion

            If nRespuesta = 2 Then
                Set rsCli = clsCli.GetPersonaCuenta(lsCtaCargo, gCapRelPersTitular)
                oSolicitud.ActualizarCapAutSinTarjetaVisto_nMovNro gsCodUser, gsCodAge, lsCtaCargo, rsCli!cperscod, nMovNroOperacion, CStr(gAhoCargoServicioRecaudo)
            End If
            Set oMovOperacion = Nothing
            nRespuesta = 0
        End If
        'CTI4 end
        
                'INICIO JHCU ENCUESTA 16-10-2019
        Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
        'FIN
    End If
    
    Exit Sub
    
Error:
    MsgBox err.Description, vbCritical, "Aviso"

End Sub

Public Function verificaPagoTrama(ByVal pnConvenio As String, ByVal pnCliente As String, ByVal pnTrama As String) As Boolean
    
    Dim paResultado As Boolean
    
    Dim oCont As COMDCaptaServicios.DCOMServicioRecaudo
    Set oCont = New COMDCaptaServicios.DCOMServicioRecaudo

    paResultado = oCont.verificaPagoTrama(pnConvenio, pnCliente, pnTrama)
    verificaPagoTrama = paResultado

End Function
'RIRO 20161102 Verifica si el/los pagos corresponden a registros mas antiguos.
Private Function ValidarFechaPagos() As String
    Dim sMensaje As String
    Dim dSeleccionados() As Seleccion
    Dim nI As Integer, nJ As Integer, P As Integer
    Dim bEncontrado As Boolean
    sMensaje = ""
    P = -1
    
    'obteniendo los datos del campo fecha de vencimiento y orden
    For nI = 1 To grdConceptoPagarVC.Rows - 1
        If grdConceptoPagarVC.TextMatrix(nI, 7) = "." Then
            P = P + 1
            ReDim Preserve dSeleccionados(P)
            dSeleccionados(P).dFechaVencimiento = grdConceptoPagarVC.TextMatrix(nI, 3)
            dSeleccionados(P).nOrdenFechaVenc = grdConceptoPagarVC.TextMatrix(nI, 9)
        End If
    Next nI
    
    'Verifica si los registros seleccionados corresponden a los mas antiguos.
    'El valor que indica el orden de las fechas, lo trae el script de la base de datos
    If P >= 0 Then
        For nI = 0 To UBound(dSeleccionados)
            bEncontrado = False
            For nJ = 0 To UBound(dSeleccionados)
                If ((nI + 1) = dSeleccionados(nJ).nOrdenFechaVenc) Then
                     bEncontrado = True
                End If
            Next nJ
            If Not bEncontrado Then
                sMensaje = "Selección incorrecta: " & vbNewLine & "Seleccione el concepto mas antiguo"
                Exit For
            End If
        Next nI
    End If
    ValidarFechaPagos = sMensaje
End Function
'END RIRO *******************************************************************


Private Function validaSVVP() As String

    Dim nI As Integer
    Dim nCont As Integer
        
    If nConvenioSeleccionado = Convenio_SV Then
        nCont = 0
        If Trim(txtNombreClienteSV.Text) = "" Then
            validaSVVP = "Verifcar datos de cliente"
            Exit Function
        ElseIf Trim(txtDOISV.Text) = "" Then
            validaSVVP = "Verifcar datos de cliente"
            Exit Function
'        ElseIf Trim(txtOtroCodigoSV.Text) = "" Then
'            validaSVVP = "Verifcar el codigo ingresado"
'            Exit Function
        ElseIf cboTipoDOISV.ListIndex = -1 Then
            validaSVVP = "Verifcar el tipo de DOI seleccionado"
            Exit Function
        End If
                
        ' Aplicando Validacion de Importe Minimos y Maximos
        Dim rsImpoteMinMax As ADODB.Recordset
        Dim nMinimo, nMaximo As Double
        Set rsImpoteMinMax = ObtenerImporteMinMax(Trim(txtCodigoBusConvenio.Text))
        
        If Not rsImpoteMinMax Is Nothing Then
            If Not rsImpoteMinMax.EOF And Not rsImpoteMinMax.BOF Then
                nMinimo = rsImpoteMinMax!nPagoMin
                nMaximo = rsImpoteMinMax!nPagoMax
            End If
        End If
        
        If CDbl(txtSubTotalSV.Text) < nMinimo Or _
           CDbl(txtSubTotalSV.Text) > nMaximo Then

            validaSVVP = " Verificar importes. Límites por operacion: Monto mínimo " & nMinimo & " - Monto máximo " & nMaximo
            Set rsImpoteMinMax = Nothing
            Exit Function

        End If
        
        For nI = 1 To grdConceptoPagarSV.Rows - 1
            nCont = nCont + 1
            If Trim(grdConceptoPagarSV.TextMatrix(nI, 1)) = "" Or _
            Val(grdConceptoPagarSV.TextMatrix(nI, 3)) <= 0 Then
                
                validaSVVP = "Verificar conceptos y montos a pagar"
                Exit Function
                
            End If
                       
        Next
        
        If nCont > 0 Then
            validaSVVP = ""
        Else
            validaSVVP = "No se registraron montos y conceptos a pagar"
        End If
        Exit Function
    ElseIf nConvenioSeleccionado = Convenio_VP Then
    
         nCont = 0
         If Trim(txtNombreClienteVX.Text) = "" Then
                validaSVVP = "Verifcar datos de cliente"
                Exit Function
        ElseIf Trim(txtDOIVX.Text) = "" Then
                validaSVVP = "Verifcar datos de cliente"
                Exit Function
'        ElseIf Trim(txtOtroCodigoVX.Text) = "" Then
'                validaSVVP = "Verifcar el codigo ingresado"
'                Exit Function
        ElseIf cboTipoDOIVX.ListIndex = -1 Then
                validaSVVP = "Verifcar el tipo de DOI seleccionado"
                Exit Function
        End If
                
        For nI = 1 To grdConceptoPagarVX.Rows - 1
        
            nCont = nCont + 1
            If Trim(grdConceptoPagarVX.TextMatrix(nI, 1)) = "" Or _
                Val(grdConceptoPagarVX.TextMatrix(nI, 3)) <= 0 Then
                
                validaSVVP = "Verificar conceptos y montos a pagar"
                Exit Function

            End If
            
        Next
        
        If nCont > 0 Then
            validaSVVP = ""
        Else
            validaSVVP = "No se registraron montos y conceptos a pagar"
        End If
        Exit Function
        
    End If
    
End Function

Private Sub registrarVP(ByVal sMovNro As String, sMovNro2 As String, _
            Optional ByVal pnITFCargoCta As Currency, _
            Optional ByVal pnMontoCargo As Currency = 0#, _
            Optional ByVal pnMovNroTransfer As Long = -1, Optional ByVal pnMonedaTrans As Moneda = gMonedaNacional, _
            Optional ByVal pnMovNroRVD As Long = 0, _
            Optional ByVal pnMontoRVD As Currency = 0)
    
    Dim cCodigoConvenio As String
    Dim cCodigoCliente As String
    Dim nTipoDOI As Integer
    Dim cDOI As String
    Dim cNombreCliente As String
    Dim dFechaCobro As Date
    Dim sBoleta As String
    Dim nI As Integer
    Dim nTamanio As Integer
    Dim arrDetalleCobro() As String 'Contiene detalle de conceptos de pago
    Dim cCadenaData As String
    Dim nDeudaActual As Double
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim nITF As Double          'Para validacion de ITF
    Dim nRedondeoITF As Double  'Para validacion de ITF
    Dim lsBoletaCargo  As String 'CTI6 ERS0112020
    Dim lsNombreClienteCargoCta As String 'CTI6 ERS0112020
    Dim MatDatosAho(14) As String 'CTI6 ERS0112020
 
    nTamanio = 0

    ' Validando si Aplica ITF
    
    nITF = Format(fgITFCalculaImpuesto(txtSubTotalVX.Text), "#,##0.00")
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(nITF))

    If nRedondeoITF > 0 Then
    
       nITF = Format(CCur(nITF) - nRedondeoITF, "#,##0.00")
       
    End If
    
    'CTI6 ERS0112020
    If Not ValidaFormaPagoVP Then Exit Sub 'CTI6 ERS0112020
    
     'Dim MatDatosAho(14) As String 'CTI6 ERS0112020
     Dim pnTipoPago As Integer
     Dim psCtaCodCargo As String
     pnMontoPagarCargo = 0#
     If CInt(Trim(Right(CmbForPagVP.Text, 10))) = gCVTipoPagoCargoCta Then
         If Mid(txtCuentaCargoVP.NroCuenta, 9, 1) = gMonedaNacional Then
             pnMontoPagarCargo = CDbl(txtTotalVX.Text)
             pnMontoCargo = pnMontoPagarCargo
         End If
         AsignaValorITFVP
     End If
     Dim clsCapN As New COMNCaptaGenerales.NCOMCaptaMovimiento
     If CInt(Trim(Right(CmbForPagVP.Text, 10))) = gCVTipoPagoCargoCta Then
        If Not clsCapN.ValidaSaldoCuenta(txtCuentaCargoVC.NroCuenta, CDbl(txtTotalVX.Text)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
            Exit Sub
        End If
     End If
    Set clsCapN = Nothing
    pnTipoPago = CInt(Trim(Right(CmbForPagVP.Text, 10)))
    psCtaCodCargo = txtCuentaCargoVP.NroCuenta
    'End

    cCodigoConvenio = Trim(txtCodigoBusConvenio.Text)
    cCodigoCliente = Trim(txtOtroCodigoVX.Text)
    nTipoDOI = Trim(Right(Trim(cboTipoDOIVX.Text), 2))
    cDOI = Trim(txtDOIVX.Text)
    cNombreCliente = Trim(txtNombreClienteVX.Text)
    dFechaCobro = gdFecSis
    
    ' Se llena un arreglo que contendra los conceptos y montos
    ' de pago correspondientes a la operacion
    For nI = 1 To grdConceptoPagarVX.Rows - 1
        
            nTamanio = nTamanio + 1
            
            ReDim Preserve arrDetalleCobro(nTamanio)
            
            'id de la Trama
            cCadenaData = grdConceptoPagarVX.TextMatrix(nI, 4) & "|"
            'Servicio
            cCadenaData = cCadenaData & "" & "|"
            'Concepto
            cCadenaData = cCadenaData & Mid(grdConceptoPagarVX.TextMatrix(nI, 1), _
                                            InStr(1, grdConceptoPagarVX.TextMatrix(nI, 1), "-") + 2, _
                                            Len(grdConceptoPagarVX.TextMatrix(nI, 1))) & "|" ' se hizo una ultima modificacion
            
            'Importe
            cCadenaData = cCadenaData & CDbl(grdConceptoPagarVX.TextMatrix(nI, 3)) & "|"
            nDeudaActual = 0#
            'Deuda Actual
            cCadenaData = cCadenaData & CDbl(nDeudaActual) & "|"
            'Monto Cobro
            cCadenaData = cCadenaData & CDbl(grdConceptoPagarVX.TextMatrix(nI, 3)) & "|"
            'Estado
            cCadenaData = cCadenaData & Pagado & "|"
            'Mora
            cCadenaData = cCadenaData & "0.00|"
            'Fecha Vencimiento
            'cCadenaData = cCadenaData & Format(CDate(gdFecSis), "yyyyMMdd") & "|"
            cCadenaData = cCadenaData & "|"
            
            arrDetalleCobro(nTamanio - 1) = cCadenaData

    Next
    
    Dim ClsServicioRecaudo As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set ClsServicioRecaudo = New COMNCaptaGenerales.NCOMCaptaMovimiento
    
    If ClsServicioRecaudo.RegistrarCobroServicioConvenio(arrDetalleCobro, cCodigoConvenio, _
                                                         cCodigoCliente, nTipoDOI, cDOI, cNombreCliente, _
                                                         dFechaCobro, sMovNro, sMovNro2, nComisionEmpresa, _
                                                         Trim(txtComisionVX.Text), nMoneda, strCuenta, nITF, , , , , _
                                                         pnTipoPago, psCtaCodCargo, MatDatosAho, gITF.gbITFAplica, pnITFCargoCta, pnMontoCargo, pnMovNroTransfer, pnMonedaTrans, pnMovNroRVD, pnMontoRVD) Then
        
        MsgBox "La operacion se realizo correctamente", vbInformation, "Aviso"
        
        Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
        sBoleta = ClsServicioRecaudo.ImprimeVaucherRecaudo(ClsServicioRecaudo.getDatosVaucherRecaudo(clsCap.GetnMovNro(sMovNro)), gbImpTMU, gsCodUser)
        
'        'CTI6 ERS0112020
'        Dim oImp As COMNContabilidad.NCOMContImprimir
'        Set oImp = New COMNContabilidad.NCOMContImprimir
'        If CInt(Trim(Right(CmbForPagVP.Text, 10))) = gCVTipoPagoCargoCta Then
'            lsBoletaCargo = oImp.ImprimeBoletaAhorro("RETIRO AHORROS", "SERVICIO DE RECAUDO", "", CStr(CDbl(pnMontoCargo) + pnITFCargoCta), lsNombreClienteCargoCta, txtCuentaCargoVP.NroCuenta, "", CDbl(MatDatosAho(10)), CDbl(MatDatosAho(3)), "", 1, CDbl(MatDatosAho(11)), , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, gsCodCMAC, 0)
'        End If
'        'CTI6 END
        
        limpiaDetalle
        'limipiarCabecera
        txtNombreClienteVX.SetFocus
        
        Do
         clsprevio.PrintSpool sLpt, sBoleta '& lsBoletaCargo 'CTI6 ERS0112020
         
        Loop While MsgBox("Desea Reimprimir el voucher?", vbQuestion + vbYesNo, "Aviso") = vbYes
        
        Set clsCap = Nothing
        
        CmbForPagVP.ListIndex = -1 'CTI6 ERS0112020
        txtCuentaCargoVP.NroCuenta = "" 'CTI6 ERS0112020
        LblNumDocVP.Caption = "" 'CTI6 ERS0112020
        pnMoneda = 0 'CTI6 ERS0112020
        IniciarVouher 'CTI7 OPEv2

        
    Else
        MsgBox "No se pudo realizar en cobro del recaudo", vbExclamation, "Aviso"
    End If
    
End Sub

Private Sub registrarSV(ByVal sMovNro As String, sMovNro2 As String, _
            Optional ByVal pnITFCargoCta As Currency, _
            Optional ByVal pnMontoCargo As Currency = 0#, _
            Optional ByVal pnMovNroTransfer As Long = -1, Optional ByVal pnMonedaTrans As Moneda = gMonedaNacional, _
            Optional ByVal pnMovNroRVD As Long = 0, _
            Optional ByVal pnMontoRVD As Currency = 0)

    Dim cCodigoConvenio As String
    Dim cCodigoCliente As String
    Dim nTipoDOI As Integer
    Dim cDOI As String
    Dim cNombreCliente As String
    Dim dFechaCobro As Date
    Dim sBoleta As String
    Dim nI As Integer
    Dim nTamanio As Integer
    Dim arrDetalleCobro() As String
    Dim cCadenaData As String
    Dim nDeudaActual As Double
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim nITF As Double          ' Para Validar ITF
    Dim nRedondeoITF As Double  ' Para Validar ITF
    Dim lsBoletaCargo  As String 'CTI6 ERS0112020
    Dim lsNombreClienteCargoCta As String 'CTI6 ERS0112020
    Dim MatDatosAho(14) As String 'CTI6 ERS0112020
    
    On Error GoTo ErrGraba
    
    ' Validando si Aplica ITF
    nITF = Format(fgITFCalculaImpuesto(txtSubTotalSV.Text), "#,##0.00")
    
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(nITF))

    If nRedondeoITF > 0 Then
    
       nITF = Format(CCur(nITF) - nRedondeoITF, "#,##0.00")
       
    End If
    
    'CTI6 ERS0112020
    If Not ValidaFormaPagoSV Then Exit Sub 'CTI6 ERS0112020
    
     'Dim MatDatosAho(14) As String 'CTI6 ERS0112020
     Dim pnTipoPago As Integer
     Dim psCtaCodCargo As String
     pnMontoPagarCargo = 0#
     If CInt(Trim(Right(CmbForPagSV.Text, 10))) = gCVTipoPagoCargoCta Then
         If Mid(txtCuentaCargoSV.NroCuenta, 9, 1) = gMonedaNacional Then
             pnMontoPagarCargo = CDbl(txtTotalSV.Text)
             pnMontoCargo = pnMontoPagarCargo
             
         End If
         AsignaValorITFSV
     End If
     Dim clsCapN As New COMNCaptaGenerales.NCOMCaptaMovimiento
     If CInt(Trim(Right(CmbForPagSV.Text, 10))) = gCVTipoPagoCargoCta Then
        If Not clsCapN.ValidaSaldoCuenta(txtCuentaCargoSV.NroCuenta, CDbl(txtTotalSV.Text)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
            Exit Sub
        End If
     End If
    Set clsCapN = Nothing
    pnTipoPago = CInt(Trim(Right(CmbForPagSV.Text, 10)))
    psCtaCodCargo = txtCuentaCargoSV.NroCuenta
    'End
            
    cCodigoConvenio = Trim(txtCodigoBusConvenio.Text)
    cCodigoCliente = Trim(txtOtroCodigoSV.Text)
    nTipoDOI = Trim(Right(Trim(cboTipoDOISV.Text), 2))
    cDOI = Trim(txtDOISV.Text)
    cNombreCliente = Trim(txtNombreClienteSV.Text)
    dFechaCobro = gdFecSis
     
    nTamanio = 0
    
    ' Se llena un arreglo que contendra los conceptos y montos
    ' de pago correspondientes a la operacion
    For nI = 1 To grdConceptoPagarSV.Rows - 1
        
        nTamanio = nTamanio + 1
        
        ReDim Preserve arrDetalleCobro(nTamanio)
        'id de la Trama
        cCadenaData = "" & "|"
        'Servicio
        cCadenaData = cCadenaData & "" & "|"
        'Concepto
        cCadenaData = cCadenaData & grdConceptoPagarSV.TextMatrix(nI, 1) & "|"
        'Importe
        cCadenaData = cCadenaData & CDbl(grdConceptoPagarSV.TextMatrix(nI, 3)) & "|"
        'Deuda Actual
        cCadenaData = cCadenaData & "0.00|"
        'Monto Cobro
        cCadenaData = cCadenaData & CDbl(grdConceptoPagarSV.TextMatrix(nI, 3)) & "|"
        'Estado
        cCadenaData = cCadenaData & Pagado & "|"
        'Mora
        cCadenaData = cCadenaData & "0.00|"
        'Fecha Vencimiento
        cCadenaData = cCadenaData & "|"
        
        arrDetalleCobro(nTamanio - 1) = cCadenaData
        
    Next
    
    Dim ClsServicioRecaudo As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set ClsServicioRecaudo = New COMNCaptaGenerales.NCOMCaptaMovimiento
    
    If ClsServicioRecaudo.RegistrarCobroServicioConvenio(arrDetalleCobro, cCodigoConvenio, _
                            cCodigoCliente, nTipoDOI, cDOI, cNombreCliente, dFechaCobro, sMovNro, sMovNro2, nComisionEmpresa, Trim(txtComisionSV.Text), nMoneda, strCuenta, nITF, , , , , _
                            pnTipoPago, psCtaCodCargo, MatDatosAho, gITF.gbITFAplica, pnITFCargoCta, pnMontoCargo, pnMovNroTransfer, pnMonedaTrans, pnMovNroRVD, pnMontoRVD) Then
        
        MsgBox "La operacion se realizo correctamente", vbInformation, "Aviso"
        
        Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
        sBoleta = ClsServicioRecaudo.ImprimeVaucherRecaudo(ClsServicioRecaudo.getDatosVaucherRecaudo(clsCap.GetnMovNro(sMovNro)), gbImpTMU, gsCodUser)
        
'        'CTI6 ERS0112020
'        Dim oImp As COMNContabilidad.NCOMContImprimir
'        Set oImp = New COMNContabilidad.NCOMContImprimir
'        If CInt(Trim(Right(CmbForPagSV.Text, 10))) = gCVTipoPagoCargoCta Then
'            lsBoletaCargo = oImp.ImprimeBoletaAhorro("RETIRO AHORROS", "SERVICIO DE RECAUDO", "", CStr(CDbl(pnMontoCargo) + pnITFCargoCta), lsNombreClienteCargoCta, txtCuentaCargoSV.NroCuenta, "", CDbl(MatDatosAho(10)), CDbl(MatDatosAho(3)), "", 1, CDbl(MatDatosAho(11)), , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, gsCodCMAC, 0)
'        End If
'        'END CTI6
        
        limpiaDetalle
        'limipiarCabecera
        txtNombreClienteSV.SetFocus
        
        Do
            clsprevio.PrintSpool sLpt, sBoleta '& lsBoletaCargo
            
        Loop While MsgBox("Desea Reimprimir el voucher?", vbQuestion + vbYesNo, "Aviso") = vbYes
        
        Set clsCap = Nothing
        CmbForPagSV.ListIndex = -1 'CTI6 ERS0112020
        txtCuentaCargoSV.NroCuenta = "" 'CTI6 ERS0112020
        LblNumDocSV.Caption = "" 'CTI6 ERS0112020
        pnMoneda = 0 'CTI6 ERS0112020

        IniciarVouher 'CTI7 OPEv2

        
    Else
    
        MsgBox "No se pudo realizar en cobro del recaudo", vbExclamation, "Aviso"
        
    End If
    
    Exit Sub
    
ErrGraba:
    Set clsCap = Nothing
    MsgBox err.Description, vbCritical, "Aviso"
     
End Sub

Private Sub registrarVI(ByVal sMovNro As String, sMovNro2 As String, _
            Optional ByVal pnITFCargoCta As Currency, _
            Optional ByVal pnMontoCargo As Currency = 0#, _
            Optional ByVal pnMovNroTransfer As Long = -1, Optional ByVal pnMonedaTrans As Moneda = gMonedaNacional, _
            Optional ByVal pnMovNroRVD As Long = 0, _
            Optional ByVal pnMontoRVD As Currency = 0)
    
    Dim cCodigoConvenio As String
    Dim cCodigoCliente As String
    Dim nTipoDOI As Integer
    Dim cDOI As String
    Dim cNombreCliente As String
    Dim dFechaCobro As Date
    Dim sBoleta As String
    Dim nI As Integer
    Dim nTamanio As Integer
    Dim arrDetalleCobro() As String
    Dim cCadenaData As String
    Dim nDeudaActual As Double
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim nITF As Double          ' Validar ITF
    Dim nRedondeoITF As Double  ' Validar ITF
    Dim lsBoletaCargo  As String 'CTI6 ERS0112020
    Dim lsNombreClienteCargoCta As String 'CTI6 ERS0112020
    Dim MatDatosAho(14) As String 'CTI6 ERS0112020
    
    nTamanio = 0
           
    ' Validando si Aplica ITF
    nITF = Format(fgITFCalculaImpuesto(txtSubTotalVI.Text), "#,##0.00")
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(nITF))
    If nRedondeoITF > 0 Then
       nITF = Format(CCur(nITF) - nRedondeoITF, "#,##0.00")
    End If
    
    'CTI6 ERS0112020
    If Not ValidaFormaPagoVI Then Exit Sub 'CTI6 ERS0112020
    
     'Dim MatDatosAho(14) As String 'CTI6 ERS0112020
     Dim pnTipoPago As Integer
     Dim psCtaCodCargo As String
     pnMontoPagarCargo = 0#
     If CInt(Trim(Right(CmbForPagVI.Text, 10))) = gCVTipoPagoCargoCta Then
         If Mid(txtCuentaCargoVI.NroCuenta, 9, 1) = gMonedaNacional Then
             pnMontoPagarCargo = CDbl(txtTotalVI.Text)
             pnMontoCargo = pnMontoPagarCargo
         End If
         AsignaValorITFVI
     End If
     Dim clsCapN As New COMNCaptaGenerales.NCOMCaptaMovimiento
     If CInt(Trim(Right(CmbForPagVI.Text, 10))) = gCVTipoPagoCargoCta Then
        If Not clsCapN.ValidaSaldoCuenta(txtCuentaCargoVC.NroCuenta, CDbl(txtTotalVI.Text)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
            Exit Sub
        End If
     End If
     Set clsCapN = Nothing
    pnTipoPago = CInt(Trim(Right(CmbForPagVI.Text, 10)))
    psCtaCodCargo = txtCuentaCargoVI.NroCuenta
    'End
                
    cCodigoConvenio = Trim(txtCodigoBusConvenio.Text)
    cCodigoCliente = Trim(txtCodigoIDVI.Text)
    nTipoDOI = rsUsuarioRecaudo!nTipoDOI
    cDOI = Trim(txtDOIVI.Text)
    cNombreCliente = Trim(txtNombreClienteVI.Text)
    dFechaCobro = gdFecSis
    
    ' Se llena un arreglo que contendra los conceptos y montos
    ' de pago correspondientes a la operacion
    For nI = 1 To grdConceptoPagarVI.Rows - 1
        If grdConceptoPagarVI.TextMatrix(nI, 4) = "." Then
            nTamanio = nTamanio + 1
            ReDim Preserve arrDetalleCobro(nTamanio)
            'Id Trama
            cCadenaData = grdConceptoPagarVI.TextMatrix(nI, 0) & "|"
            'Servicio
            cCadenaData = cCadenaData & grdConceptoPagarVI.TextMatrix(nI, 1) & "|"
            'Concepto
            cCadenaData = cCadenaData & grdConceptoPagarVI.TextMatrix(nI, 2) & "|"
            'Importe
            cCadenaData = cCadenaData & CDbl(rsUsuarioRecaudo!nDeudaActual) & "|"
            nDeudaActual = CDbl(rsUsuarioRecaudo!nDeudaActual) - CDbl(grdConceptoPagarVI.TextMatrix(nI, 5))
            'Deuda Actual
            cCadenaData = cCadenaData & nDeudaActual & "|"
            'Monto Cobro
            cCadenaData = cCadenaData & CDbl(grdConceptoPagarVI.TextMatrix(nI, 5)) & "|"
            'Estado
            If nDeudaActual <= 0 Then
                cCadenaData = cCadenaData & Pagado & "|"
            Else
                cCadenaData = cCadenaData & Pagando & "|"
            End If
            'Mora
            cCadenaData = cCadenaData & "0.00|"
            'Fecha Vencimiento
            cCadenaData = cCadenaData & "|"
            
            arrDetalleCobro(nTamanio - 1) = cCadenaData
        
        End If
        rsUsuarioRecaudo.MoveNext
        
    Next
    rsUsuarioRecaudo.MoveFirst
    If cCadenaData = "" Then
    
        MsgBox "No se seleccionó ningún concepto de pago", vbExclamation, "Aviso"
        Exit Sub
        
    End If
    
    Dim ClsServicioRecaudo As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set ClsServicioRecaudo = New COMNCaptaGenerales.NCOMCaptaMovimiento
    
    If ClsServicioRecaudo.RegistrarCobroServicioConvenio(arrDetalleCobro, cCodigoConvenio, cCodigoCliente, _
                                                        nTipoDOI, cDOI, cNombreCliente, dFechaCobro, _
                                                        sMovNro, sMovNro2, nComisionEmpresa, Trim(txtComisionVI.Text), nMoneda, strCuenta, nITF, , , , , _
                                                        pnTipoPago, psCtaCodCargo, MatDatosAho, gITF.gbITFAplica, pnITFCargoCta, pnMontoCargo, pnMovNroTransfer, pnMonedaTrans, pnMovNroRVD, pnMontoRVD) Then
    
        MsgBox "La operacion se realizo correctamente", vbInformation, "Aviso"
                
        Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
        sBoleta = ClsServicioRecaudo.ImprimeVaucherRecaudo(ClsServicioRecaudo.getDatosVaucherRecaudo(clsCap.GetnMovNro(sMovNro)), gbImpTMU, gsCodUser)
        
'        'CTI6 ERS0112020
'        Dim oImp As COMNContabilidad.NCOMContImprimir
'        Set oImp = New COMNContabilidad.NCOMContImprimir
'        If CInt(Trim(Right(CmbForPagVI.Text, 10))) = gCVTipoPagoCargoCta Then
'            lsBoletaCargo = oImp.ImprimeBoletaAhorro("RETIRO AHORROS", "SERVICIO DE RECAUDO", "", CStr(CDbl(pnMontoCargo) + pnITFCargoCta), lsNombreClienteCargoCta, txtCuentaCargoVI.NroCuenta, "", CDbl(MatDatosAho(10)), CDbl(MatDatosAho(3)), "", 1, CDbl(MatDatosAho(11)), , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, gsCodCMAC, 0)
'        End If
'        'CTI6 END
        
        limpiaDetalle
        'limipiarCabecera
        cmdBuscarPersonaVI.SetFocus
        Do
            clsprevio.PrintSpool sLpt, sBoleta '& lsBoletaCargo
        Loop While MsgBox("Desea Reimprimir el voucher?", vbQuestion + vbYesNo, "Aviso") = vbYes
        Set clsCap = Nothing
        CmbForPagVI.ListIndex = -1 'CTI6 ERS0112020
        txtCuentaCargoVI.NroCuenta = "" 'CTI6 ERS0112020
        LblNumDocVI.Caption = "" 'CTI6 ERS0112020
        pnMoneda = 0 'CTI6 ERS0112020
        IniciarVouher 'CTI7 OPEv2

    Else
        MsgBox "No se pudo realizar en cobro del recaudo", vbExclamation, "Aviso"
    End If
    
End Sub
Private Sub registrarVC(ByVal sMovNro As String, sMovNro2 As String, _
            Optional ByVal pnITFCargoCta As Currency, _
            Optional ByVal pnMontoCargo As Currency = 0#, _
            Optional ByVal pnMovNroTransfer As Long = -1, Optional ByVal pnMonedaTrans As Moneda = gMonedaNacional, _
            Optional ByVal pnMovNroRVD As Long = 0, _
            Optional ByVal pnMontoRVD As Currency = 0)
    
    Dim cCodigoConvenio As String
    Dim cCodigoCliente As String
    Dim nTipoDOI As Integer
    Dim cDOI As String
    Dim cNombreCliente As String
    Dim dFechaCobro As Date
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim sBoleta As String
    Dim nITF As Double          ' Validar ITF
    Dim nRedondeoITF As Double  ' Validar ITF
    Dim cUrlSimaynas As String  'CTI1 ERS027-2019
    Dim lsBoletaCargo  As String 'CTI6 ERS0112020
    Dim lsNombreClienteCargoCta As String 'CTI6 ERS0112020
    Dim MatDatosAho(14) As String 'CTI6 ERS0112020
    
    'Validando si Aplica ITF
            
    nITF = Format(fgITFCalculaImpuesto(txtSubTotalVC.Text), "#,##0.00")
    nRedondeoITF = fgDiferenciaRedondeoITF(CCur(nITF))
    If nRedondeoITF > 0 Then
       nITF = Format(CCur(nITF) - nRedondeoITF, "#,##0.00")
    End If
    
    'Datos Unicos -----------------------------------------
    cCodigoConvenio = Trim(txtCodigoBusConvenio.Text)
    cCodigoCliente = Trim(txtCodigoIDVC.Text)
    nTipoDOI = rsUsuarioRecaudo!nTipoDOI
    cDOI = Trim(txtDOIVC.Text)
    cNombreCliente = Trim(txtNombreClienteVC.Text)
    dFechaCobro = gdFecSis
    ' Datos Unicos -----------------------------------------
    
    'CTI6 ERS0112020
    If Not ValidaFormaPagoVC Then Exit Sub 'CTI6 ERS0112020
    

     Dim pnTipoPago As Integer
     Dim psCtaCodCargo As String
     pnMontoPagarCargo = 0#
     If CInt(Trim(Right(CmbForPagVC.Text, 10))) = gCVTipoPagoCargoCta Then
         If Mid(txtCuentaCargoVC.NroCuenta, 9, 1) = gMonedaNacional Then
             pnMontoPagarCargo = CDbl(txtTotalVC.Text)
             pnMontoCargo = pnMontoPagarCargo
         End If
         AsignaValorITFVC
     End If
     Dim clsCapN As New COMNCaptaGenerales.NCOMCaptaMovimiento
     If CInt(Trim(Right(CmbForPagVC.Text, 10))) = gCVTipoPagoCargoCta Then
        If Not clsCapN.ValidaSaldoCuenta(txtCuentaCargoVC.NroCuenta, CDbl(txtTotalVC.Text)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
            Exit Sub
        End If
     End If
     Set clsCapN = Nothing
    pnTipoPago = CInt(Trim(Right(CmbForPagVC.Text, 10)))
    psCtaCodCargo = txtCuentaCargoVC.NroCuenta
    'End
    
    Dim nI As Integer
    Dim nTamanio As Integer
    Dim arrDetalleCobro() As String
    Dim cCadenaData As String

    nConvenio = cCodigoConvenio 'ADD BY PTI1 20210723
    nCliente = cCodigoCliente 'ADD BY PTI1 20210723
    nTamanio = 0

    For nI = 1 To grdConceptoPagarVC.Rows - 1
        'If grdConceptoPagarVC.TextMatrix(nI, 5) = "." Then
        If grdConceptoPagarVC.TextMatrix(nI, 7) = "." Then
            nTrama = Trim(grdConceptoPagarVC.TextMatrix(nI, 0)) 'ADD BY PTI1 20210723
            If verificaPagoTrama(nConvenio, nCliente, nTrama) Then 'ADD BY PTI1 20210723
                MsgBox "La cuota de " & grdConceptoPagarVC.TextMatrix(nI, 2) & " ya fue cancelada", vbCritical, "Aviso" 'ADD BY PTI1 20210723
                limpiaDetalle
                'limipiarCabecera
                cmdBuscarPersonaVC.SetFocus
                Exit Sub
            Else
                nTamanio = nTamanio + 1
                ReDim Preserve arrDetalleCobro(nTamanio)
                'ID de la Trama
                cCadenaData = Trim(grdConceptoPagarVC.TextMatrix(nI, 0)) & "|"
                'Servicio
                cCadenaData = cCadenaData & Replace(Trim(grdConceptoPagarVC.TextMatrix(nI, 1)), ".", "") & "|"
                'Concepto
                cCadenaData = cCadenaData & Trim(grdConceptoPagarVC.TextMatrix(nI, 2)) & "|"
                'Importe
                cCadenaData = cCadenaData & CDbl(grdConceptoPagarVC.TextMatrix(nI, 5)) & "|"
                'Deuda Actual
                cCadenaData = cCadenaData & "0.00" & "|"
                'Monto Cobro
                cCadenaData = cCadenaData & CDbl(grdConceptoPagarVC.TextMatrix(nI, 5)) & "|"
                'Estado
                cCadenaData = cCadenaData & Pagado & "|"
                'Mora
                cCadenaData = cCadenaData & CDbl(grdConceptoPagarVC.TextMatrix(nI, 6)) & "|"
                'Fecha Vencimiento
                If nConvenioSeleccionado = Convenio_VCM Or verificaWS = True Then 'CTI1 ERS027-2019
                    cCadenaData = cCadenaData & Format(CDate(grdConceptoPagarVC.TextMatrix(nI, 3)), "yyyyMMdd") & "|"
                Else
                    cCadenaData = cCadenaData & "|"
                End If
                arrDetalleCobro(nTamanio - 1) = cCadenaData
            End If

        End If
    Next

    If cCadenaData = "" Then
        MsgBox "no se seleccionó ningún concepto de pago", vbExclamation, "Aviso"
        Exit Sub
    End If

    Dim ClsServicioRecaudo As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set ClsServicioRecaudo = New COMNCaptaGenerales.NCOMCaptaMovimiento
    
    If verificaWS = True Then 'CTI1 ERS027-2019
        cUrlSimaynas = Trim(LeeConstanteSist(708))
    End If 'CTI1 ERS027-2019

    If ClsServicioRecaudo.RegistrarCobroServicioConvenio(arrDetalleCobro, _
                        cCodigoConvenio, cCodigoCliente, nTipoDOI, cDOI, cNombreCliente, _
                        dFechaCobro, sMovNro, sMovNro2, nComisionEmpresa, Trim(txtComisionVC.Text), _
                        nMoneda, strCuenta, nITF, , , verificaWS, cUrlSimaynas, pnTipoPago, psCtaCodCargo, _
                        MatDatosAho, gITF.gbITFAplica, pnITFCargoCta, pnMontoCargo, pnMovNroTransfer, pnMonedaTrans, pnMovNroRVD, pnMontoRVD) Then
        'CTI1 ERS027-2019 add verificaWS, cUrlSimaynas

        MsgBox "La operacion se realizo correctamente", vbInformation, "Aviso"

        Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
        sBoleta = ClsServicioRecaudo.ImprimeVaucherRecaudo(ClsServicioRecaudo.getDatosVaucherRecaudo(clsCap.GetnMovNro(sMovNro)), gbImpTMU, gsCodUser)

'        'CTI6 ERS0112020
'        Dim oImp As COMNContabilidad.NCOMContImprimir
'        Set oImp = New COMNContabilidad.NCOMContImprimir
'        If CInt(Trim(Right(CmbForPagVC.Text, 10))) = gCVTipoPagoCargoCta Then
'            lsBoletaCargo = oImp.ImprimeBoletaAhorro("RETIRO AHORROS", "SERVICIO DE RECAUDO", "", CStr(CDbl(pnMontoCargo) + pnITFCargoCta), lsNombreClienteCargoCta, txtCuentaCargoVC.NroCuenta, "", CDbl(MatDatosAho(10)), CDbl(MatDatosAho(3)), "", 1, CDbl(MatDatosAho(11)), , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, gsCodCMAC, 0)
'        End If
'        'END CTI6
        
        limpiaDetalle
        'limipiarCabecera
        cmdBuscarPersonaVC.SetFocus

        Do
            clsprevio.PrintSpool sLpt, sBoleta '& lsBoletaCargo
        Loop While MsgBox("Desea Reimprimir el voucher?", vbInformation + vbYesNo, "Aviso") = vbYes
        Set clsCap = Nothing
        
        CmbForPagVC.ListIndex = -1 'CTI6 ERS0112020
        txtCuentaCargoVC.NroCuenta = "" 'CTI6 ERS0112020
        LblNumDocVC.Caption = "" 'CTI6 ERS0112020
        pnMoneda = 0 'CTI6 ERS0112020
        IniciarVouher 'CTI7 OPEv2
        
    Else
        MsgBox "No se pudo realizar en cobro del recaudo", vbExclamation, "Aviso"
    End If
           
End Sub

Private Sub cmdQuitarSV_Click()
    grdConceptoPagarSV.EliminaFila grdConceptoPagarSV.row
    grdConceptoPagarSV_OnCellChange 0, 0
End Sub

Private Sub cmdQuitarSV_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    cmdGuardar.SetFocus
    End If

End Sub

Private Sub cmdQuitarVX_Click()
      grdConceptoPagarVX.EliminaFila grdConceptoPagarVX.row
      grdConceptoPagarVX_OnChangeCombo
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub vistaPestana(ByVal strConvenio As String)
    Dim cTipoConvenio As String
    cTipoConvenio = Mid(Trim(strConvenio), 14, 2)
    stContenedorValidacion.Enabled = True
    stContenedorValidacion.TabVisible(0) = False
    stContenedorValidacion.TabVisible(1) = False
    stContenedorValidacion.TabVisible(2) = False
    stContenedorValidacion.TabVisible(3) = False
        
    If cTipoConvenio = "VC" Then
        stContenedorValidacion.TabVisible(0) = True
        stContenedorValidacion.Tab = 0
                
        
        If Mid(Trim(strConvenio), 16, 1) <> "M" Then
            'CONVENIO VC
            grdConceptoPagarVC.ColWidth(1) = 1905
            grdConceptoPagarVC.ColWidth(2) = 1905
            grdConceptoPagarVC.ColWidth(3) = 0
            grdConceptoPagarVC.ColWidth(4) = 1200
            grdConceptoPagarVC.ColWidth(5) = 1305
            grdConceptoPagarVC.ColWidth(6) = 0
            grdConceptoPagarVC.ColWidth(7) = 1005
            stContenedorValidacion.TabCaption(0) = "Validación Completa"
            nConvenioSeleccionado = Convenio_VC
        Else
            'CONVENIO MYPE
            grdConceptoPagarVC.ColWidth(1) = 1000
            grdConceptoPagarVC.ColWidth(2) = 1100
            grdConceptoPagarVC.ColWidth(3) = 1200
            grdConceptoPagarVC.ColWidth(4) = 850
            grdConceptoPagarVC.ColWidth(5) = 1100
            grdConceptoPagarVC.ColWidth(6) = 1100
            grdConceptoPagarVC.ColWidth(7) = 650
            stContenedorValidacion.TabCaption(0) = "Valid. Comp. MYPE"
            nConvenioSeleccionado = Convenio_VCM
        End If
                        
        cmdBuscarPersonaVC.SetFocus
        
    ElseIf cTipoConvenio = "VI" Then
        stContenedorValidacion.TabVisible(1) = True
        stContenedorValidacion.Tab = 1
        nConvenioSeleccionado = Convenio_VI
        cmdBuscarPersonaVI.SetFocus
        
    ElseIf cTipoConvenio = "SV" Then
        stContenedorValidacion.TabVisible(2) = True
        stContenedorValidacion.Tab = 2
        nConvenioSeleccionado = Convenio_SV
        cboTipoDOISV.ListIndex = 0
        txtNombreClienteSV.SetFocus
        
    ElseIf cTipoConvenio = "VP" Then
        stContenedorValidacion.TabVisible(3) = True
        stContenedorValidacion.Tab = 3
        nConvenioSeleccionado = Convenio_VP
        cboTipoDOIVX.ListIndex = 0
        txtNombreClienteVX.SetFocus
        
    End If
    
End Sub

Private Sub Form_Load()

    stContenedorValidacion.TabVisible(0) = True
    stContenedorValidacion.TabVisible(1) = False
    stContenedorValidacion.TabVisible(2) = False
    stContenedorValidacion.TabVisible(3) = False
    lblDescripcionCodigoVC.Caption = ""
    stContenedorValidacion.Enabled = False
    cmdGuardar.Enabled = False
    cmdcancelar.Enabled = False
    bFocusGrid = False
    cboTransferMonedaVC.ListIndex = -1
    cboTransferMonedaVI.ListIndex = -1
    cboTransferMonedaSV.ListIndex = -1
    cboTransferMonedaVP.ListIndex = -1
    Call CargaControlesVC 'CTI6 ERS0112020
    Call CargaControlesSV 'CTI6 ERS0112020
    Call CargaControlesVI 'CTI6 ERS0112020
    Call CargaControlesVP 'CTI6 ERS0112020
    'CTI7 OPEv2******************************************************
    Me.lblTTCCDVC.Caption = Format(gnTipCambioC, "#,#0.0000")
    Me.lblTTCVDVC.Caption = Format(gnTipCambioV, "#,#0.0000")
    
    Me.lblTTCCDVI.Caption = Format(gnTipCambioC, "#,#0.0000")
    Me.lblTTCVDVI.Caption = Format(gnTipCambioV, "#,#0.0000")
    
    Me.lblTTCCDSV.Caption = Format(gnTipCambioC, "#,#0.0000")
    Me.lblTTCVDSV.Caption = Format(gnTipCambioV, "#,#0.0000")
    
    Me.lblTTCCDVP.Caption = Format(gnTipCambioC, "#,#0.0000")
    Me.lblTTCVDVP.Caption = Format(gnTipCambioV, "#,#0.0000")

    '****************************************************************
    Set loVistoElectronico = New frmVistoElectronico 'CTI4 ERS0112020
End Sub

Private Sub grdConceptoPagarSV_GotFocus()
    bFocusGrid = True
End Sub

Private Sub grdConceptoPagarSV_LostFocus()
    bFocusGrid = False
End Sub

Private Sub grdConceptoPagarSV_OnCellChange(pnRow As Long, pnCol As Long)
    grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 1) = UCase(grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 1))
    Dim nSubTotal As Double
    Dim nI As Integer
    If grdConceptoPagarSV.Col = 3 Then
        If Not IsNumeric(grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 3)) Then
            MsgBox "Valor de la celda no es numerico", vbExclamation, "Aviso"
            grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 3) = "0.00"
        Else
            If grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 3) < 0 Then
             MsgBox "Valor de la celda es menor que 0.00", vbExclamation, "Aviso"
             grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 3) = "0.00"
            End If
        End If
    End If
    For nI = 1 To grdConceptoPagarSV.Rows - 1
        If grdConceptoPagarSV.TextMatrix(nI, 3) <> "" Then
            nSubTotal = nSubTotal + CDbl(IIf(Not IsNumeric(grdConceptoPagarSV.TextMatrix(nI, 3)), 0, grdConceptoPagarSV.TextMatrix(nI, 3)))
        End If
    Next
    
    txtSubTotalSV.Text = Format(nSubTotal, "#,##0.00")
    'calculoComision grdConceptoPagarSV, txtComisionSV, -1, txtSubTotalSV
    
    calculoComisionNew 'RIRO20150320
    
    ' Suma y muesta el Total
    Dim strValor As Double
    strValor = Math.Round((CDbl(txtSubTotalSV.Text) + CDbl(txtComisionSV.Text)), 2)
    txtTotalSV.Text = Format(strValor, "#,##0.00")
    
    If grdConceptoPagarSV.Col = 1 And _
       Trim(grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 3)) = "" Then
        
        grdConceptoPagarSV.TextMatrix(grdConceptoPagarSV.row, 3) = "0.00"
        
    End If
     
End Sub

Private Sub grdConceptoPagarSV_OnRowChange(pnRow As Long, pnCol As Long)

    If pnRow = grdConceptoPagarSV.Rows - 1 And pnCol = 3 Then
    
          DoEvents
          cmdAgregarSV.SetFocus
          
    End If

End Sub

Private Sub grdConceptoPagarSV_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)

    Dim sColumnas() As String
    sColumnas = Split(grdConceptoPagarSV.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "No es posible editar este campo", vbInformation, "Aviso"
    End If
    
End Sub

Private Sub grdConceptoPagarVC_GotFocus()
    bFocusGrid = True
End Sub

Private Sub grdConceptoPagarVC_LostFocus()
    bFocusGrid = False
End Sub

Private Sub grdConceptoPagarVC_KeyPress(KeyAscii As Integer)

    If grdConceptoPagarVC.Col = 4 Then
        KeyAscii = 13
        Exit Sub
    End If

End Sub
Private Sub grdConceptoPagarVC_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    
    Dim nSubTotal As Double
    Dim nI As Integer
    
    Dim nColImp As Double ' columna importe
    Dim nColMor As Double ' columna mora
    
    nColImp = 5
    nColMor = 6
    
    For nI = 1 To grdConceptoPagarVC.Rows - 1
        If grdConceptoPagarVC.TextMatrix(nI, 7) = "." Then
            'Sumando Importe
            nSubTotal = nSubTotal + CDbl(IIf(Not IsNumeric(grdConceptoPagarVC.TextMatrix(nI, nColImp)), 0, grdConceptoPagarVC.TextMatrix(nI, nColImp)))
            'Sumando Mora
            nSubTotal = nSubTotal + CDbl(IIf(Not IsNumeric(grdConceptoPagarVC.TextMatrix(nI, nColMor)), 0, grdConceptoPagarVC.TextMatrix(nI, nColMor)))
        End If
    Next
    txtSubTotalVC.Text = Format(nSubTotal, "#,##00.00")
    'calculoComision grdConceptoPagarVC, txtComisionVC, 7, txtSubTotalVC
    calculoComisionNew 'RIRO20150320
    
    ' Suma y muesta el Total
    Dim strValor As Double
    strValor = Math.Round((CDbl(txtSubTotalVC.Text) + CDbl(txtComisionVC.Text)), 2)
    txtTotalVC.Text = Format(strValor, "#,##0.00")
    
End Sub

'RIRO20150326 ****************
Private Sub calculoComisionNew()
        
    Dim nI As Integer
    Dim i As Integer
    Dim nColImp As Double 'columna importe
    Dim nColMor As Double 'columna mora

    Dim vLista() As Double '
    ReDim vLista(3, 0) 'columnas de arreglo vLista:1: Autoincrement, 2: importe, 3: mora
        
    Dim oServicio As New COMDCaptaServicios.DCOMServicioRecaudo
        
    If nConvenioSeleccionado = Convenio_VC Or nConvenioSeleccionado = Convenio_VCM Then
        nColImp = 5
        nColMor = 6
        For nI = 1 To grdConceptoPagarVC.Rows - 1
            If grdConceptoPagarVC.TextMatrix(nI, 7) = "." Then
                i = i + 1
                ReDim Preserve vLista(3, 0 To i)
                vLista(1, i) = i
                vLista(2, i) = CDbl(IIf(Not IsNumeric(grdConceptoPagarVC.TextMatrix(nI, nColImp)), 0, grdConceptoPagarVC.TextMatrix(nI, nColImp)))
                vLista(3, i) = CDbl(IIf(Not IsNumeric(grdConceptoPagarVC.TextMatrix(nI, nColMor)), 0, grdConceptoPagarVC.TextMatrix(nI, nColMor)))
            End If
        Next
        If UBound(vLista, 2) = 0 Then
            nComisionCliente = 0
            nComisionEmpresa = 0
            txtComisionVC.Text = "0.00"
            Exit Sub
        End If
        oServicio.CalculaComisionRecaudo Trim(txtCodigoBusConvenio.Text), Trim(txtCodigoIDVC.Text), vLista, nComisionCliente, nComisionEmpresa
        txtComisionVC.Text = Format(nComisionCliente, "#,##0.00")
        
    ElseIf nConvenioSeleccionado = Convenio_VI Then
        nColImp = 5
        nColMor = 0 ' no hay mora
        For nI = 1 To grdConceptoPagarVI.Rows - 1
            If grdConceptoPagarVI.TextMatrix(nI, 4) = "." Then
                i = i + 1
                ReDim Preserve vLista(3, 0 To i)
                vLista(1, i) = i
                vLista(2, i) = CDbl(IIf(Not IsNumeric(grdConceptoPagarVI.TextMatrix(nI, nColImp)), 0, grdConceptoPagarVI.TextMatrix(nI, nColImp)))
                vLista(3, i) = 0
            End If
        Next
        If UBound(vLista, 2) = 0 Then
            nComisionCliente = 0
            nComisionEmpresa = 0
            txtComisionVI.Text = "0.00"
            Exit Sub
        End If
        oServicio.CalculaComisionRecaudo Trim(txtCodigoBusConvenio.Text), Trim(txtCodigoIDVI.Text), vLista, nComisionCliente, nComisionEmpresa
        txtComisionVI.Text = Format(nComisionCliente, "#,##0.00")
        
    ElseIf nConvenioSeleccionado = Convenio_SV Then
        nColImp = 3
        nColMor = 0 ' no hay mora
        For nI = 1 To grdConceptoPagarSV.Rows - 1
            If Trim(grdConceptoPagarSV.TextMatrix(nI, 3)) <> "" Then
                i = i + 1
                ReDim Preserve vLista(3, 0 To i)
                vLista(1, i) = i
                vLista(2, i) = CDbl(IIf(Not IsNumeric(grdConceptoPagarSV.TextMatrix(nI, nColImp)), 0, grdConceptoPagarSV.TextMatrix(nI, nColImp)))
                vLista(3, i) = 0
            End If
        Next
        If UBound(vLista, 2) = 0 Then
            nComisionCliente = 0
            nComisionEmpresa = 0
            txtComisionSV.Text = "0.00"
            Exit Sub
        End If
        oServicio.CalculaComisionRecaudo Trim(txtCodigoBusConvenio.Text), Trim(txtOtroCodigoSV.Text), vLista, nComisionCliente, nComisionEmpresa   ' revisar
        txtComisionSV.Text = Format(nComisionCliente, "#,##0.00")
        
    ElseIf nConvenioSeleccionado = Convenio_VP Then
        nColImp = 3
        nColMor = 0 ' no hay mora
        Dim nSubTotal As Double
        For nI = 1 To grdConceptoPagarVX.Rows - 1
            If grdConceptoPagarVX.TextMatrix(nI, 3) <> "" Then
                i = i + 1 'APRI 20170502
                ReDim Preserve vLista(3, 0 To i)
                vLista(1, i) = i
                vLista(2, i) = CDbl(IIf(Not IsNumeric(grdConceptoPagarVX.TextMatrix(nI, nColImp)), 0, grdConceptoPagarVX.TextMatrix(nI, nColImp)))
                vLista(3, i) = 0
            End If
        Next
        If UBound(vLista, 2) = 0 Then
            nComisionCliente = 0
            nComisionEmpresa = 0
            txtComisionVX.Text = "0.00"
            Exit Sub
        End If
        oServicio.CalculaComisionRecaudo Trim(txtCodigoBusConvenio.Text), Trim(txtOtroCodigoVX.Text), vLista, nComisionCliente, nComisionEmpresa
        txtComisionVX.Text = Format(nComisionCliente, "#,##0.00")
        
    End If
End Sub
'END RIRO *************************

Private Sub calculoComision(ByVal flxTemp As FlexEdit, _
                            ByVal txtComision As TextBox, _
                            Optional ByVal nIndiceCheck As Integer, _
                            Optional ByVal txtSubTotal As TextBox)
    Dim nI As Integer
    Dim nCount As Integer
    Dim nMonto As Double
    Dim nnMonto As Double
    Dim nnMontoEmp As Double
    
    Dim nMontoTemporal As Double
                    
    '================================================================================================
    If nIndiceCheck = -1 Then
        Dim rsSV_VC As Recordset
        Set rsSV_VC = New Recordset
                
        Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
        Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
        Set rsSV_VC = ClsServicioRecaudo.GetBUscarConvenioXCodigo(Trim(txtCodigoBusConvenio.Text))
        Set rsUsuarioRecaudo = rsSV_VC
    End If
    '================================================================================================
    If Not rsUsuarioRecaudo Is Nothing Or nConvenioSeleccionado = Convenio_SV Then
        If rsUsuarioRecaudo!nTipoCobro = TipoCobro_porConcepto Then ' Valida Tipo Cobro X Concepto
            If rsUsuarioRecaudo!nTipoCalculo = TipoCalculo_fijo Then 'Valida Tipo Calculo -> Calculo Fijo
                
                ' Caso: Por Concepto->Fijo==============================================================={
                    
                    ' Valida Distribucion es Fija
                    If rsUsuarioRecaudo!nDistribucion = Distribucion_fijo Then

                       For nI = 1 To flxTemp.Rows - 1
                           If nIndiceCheck = -1 Then nIndiceCheck = 0
                           If flxTemp.TextMatrix(nI, nIndiceCheck) = "." Or nIndiceCheck = 0 Then
                                nCount = nCount + 1
                           End If
                       Next
                       
                       txtComision.Text = Format(nCount * rsUsuarioRecaudo!nDistCliente, "#,##0.00")
                       nComisionEmpresa = Format(nCount * rsUsuarioRecaudo!nDistEmpresa, "#,##0.00")
                       txtComisionVC.Text = Format$(nCount * rsUsuarioRecaudo!nDistCliente, "#,##0.00")
                       
                    ' Distribucio Porcentaje
                    Else
                       For nI = 1 To flxTemp.Rows - 1
                            If nIndiceCheck = -1 Then nIndiceCheck = 0
                            If flxTemp.TextMatrix(nI, nIndiceCheck) = "." Or nIndiceCheck = 0 Then
                                nCount = nCount + 1
                            End If
                       Next
                      txtComision.Text = Format(nCount * ((rsUsuarioRecaudo!nDistCliente) / 100) _
                                                   * (rsUsuarioRecaudo!nComision), "#,##0.00" _
                                                   )
                      nComisionEmpresa = Format(nCount * ((rsUsuarioRecaudo!nDistEmpresa) / 100) _
                                                   * (rsUsuarioRecaudo!nComision), "#,##0.00" _
                                                   )
                       txtComisionVC.Text = Format(nCount * ((rsUsuarioRecaudo!nDistCliente) / 100) _
                                                   * (rsUsuarioRecaudo!nComision), "#,##0.00")
                                                 
                    End If
                 ' Caso: Por Concepto->Fijo===============================================================}

            Else ' Valida Tipo Calculo -> Porcentual
                 ' Caso: Por Concepto->Porcentual==============================================================={
                    If rsUsuarioRecaudo!nDistribucion = Distribucion_porcentaje Then ' Valida Distribucion es Porcentual
                       For nI = 1 To flxTemp.Rows - 1
                       If nIndiceCheck = -1 Then nIndiceCheck = 0
                            If flxTemp.TextMatrix(nI, nIndiceCheck) = "." Or nIndiceCheck = 0 Then
                            nMontoTemporal = (((rsUsuarioRecaudo!nComision) / 100) * _
                                            IIf(Not IsNumeric(flxTemp.TextMatrix(nI, IIf(nIndiceCheck = 0, 3, IIf(nIndiceCheck = 7, 5, 5)))), 0, _
                                                              flxTemp.TextMatrix(nI, IIf(nIndiceCheck = 0, 3, IIf(nIndiceCheck = 7, 5, 5)))))
                                                              
                            'Solo si es convenio Validacion Completa MYPE
                            If Mid(Trim(txtCodigoBusConvenio.Text), 14, 3) = "VCM" Then
                                nMontoTemporal = nMontoTemporal + (((rsUsuarioRecaudo!nComision) / 100) * _
                                                                    IIf(Not IsNumeric(flxTemp.TextMatrix(nI, 6)), 0, _
                                                                                      flxTemp.TextMatrix(nI, 6)))
                            End If
                            'Fin de calculo para Convenios MYPE
                                                              
                            If nMontoTemporal < rsUsuarioRecaudo!nMinimo Then
                                nMontoTemporal = rsUsuarioRecaudo!nMinimo
                            ElseIf nMontoTemporal > rsUsuarioRecaudo!nMaximo Then
                                nMontoTemporal = rsUsuarioRecaudo!nMaximo
                            End If
                            nnMontoEmp = nnMontoEmp + nMontoTemporal * ((rsUsuarioRecaudo!nDistEmpresa) / 100)
                            nnMonto = nnMonto + nMontoTemporal * ((rsUsuarioRecaudo!nDistCliente) / 100)
                            nCount = nCount + 1 '
                           End If
                       Next
                       If nCount = 0 Then
                        txtComision.Text = "0.00"
                        txtTotalVC.Text = "0.00"
                        Exit Sub
                       End If
                       txtComision.Text = Format(nnMonto, "#,##0.00")
                       nComisionEmpresa = nnMontoEmp
                    End If
                ' Caso: Por Concepto->Porcentual===============================================================}
            End If
            
        Else 'Valida Tipo Cobro X Operacion
            If rsUsuarioRecaudo!nTipoCalculo = TipoCalculo_fijo Then 'Valida Tipo Calculo -> Calculo Fijo
                ' Caso: Por Concepto->Fijo==============================================================={
                    If rsUsuarioRecaudo!nDistribucion = Distribucion_fijo Then ' Valida Distribucion es Fija
                       For nI = 1 To flxTemp.Rows - 1
                       If nIndiceCheck = -1 Then nIndiceCheck = 0
                           If flxTemp.TextMatrix(nI, nIndiceCheck) = "." Or nIndiceCheck = 0 Then
                               nCount = nCount + 1
                           End If
                       Next
                       If nCount = 0 Then
                        txtComision.Text = "0.00"
                        txtTotalVC.Text = "0.00"
                        Exit Sub
                       End If
                       txtComision.Text = Format(rsUsuarioRecaudo!nDistCliente, "#,##0.00")
                       nComisionEmpresa = Format(rsUsuarioRecaudo!nDistEmpresa, "#,##0.00")
                    Else ' Distribucio Porcentaje
                       For nI = 1 To flxTemp.Rows - 1
                       If nIndiceCheck = -1 Then nIndiceCheck = 0
                          If flxTemp.TextMatrix(nI, nIndiceCheck) = "." Or nIndiceCheck = 0 Then
                              nCount = nCount + 1
                          End If
                       Next
                       If nCount = 0 Then
                        txtComision.Text = "0.00"
                        txtTotalVC.Text = "0.00"
                        Exit Sub
                       End If
                       txtComision.Text = Format(((rsUsuarioRecaudo!nDistCliente) / 100) _
                                                   * (rsUsuarioRecaudo!nComision), "#,##0.00" _
                                                   )
                        nComisionEmpresa = Format(((rsUsuarioRecaudo!nDistEmpresa) / 100) _
                                                   * (rsUsuarioRecaudo!nComision), "#,##0.00" _
                                                   )
                    End If
                ' Caso: Por Concepto->Fijo===============================================================}
            
            Else ' Valida Tipo Calculo -> Porcentual
               ' Caso: Por Concepto->Porcentual==============================================================={
                    If rsUsuarioRecaudo!nDistribucion = Distribucion_porcentaje Then ' Valida Distribucion es Porcentual
                                           
                       For nI = 1 To flxTemp.Rows - 1
                       If nIndiceCheck = -1 Then nIndiceCheck = 0
                           If flxTemp.TextMatrix(nI, nIndiceCheck) = "." Or nIndiceCheck = 0 Then
                               nCount = nCount + 1
                           End If
                       Next
                       If nCount = 0 Then
                        txtComision.Text = "0.00"
                        txtTotalVC.Text = "0.00"
                        Exit Sub
                       End If
                       'Dim nMonto As Double
                       
                       nMontoTemporal = (CDbl(txtSubTotal.Text) * ((rsUsuarioRecaudo!nComision) / 100))
                       
                       If nMontoTemporal < rsUsuarioRecaudo!nMinimo Then
                           nMontoTemporal = rsUsuarioRecaudo!nMinimo
                       ElseIf nMontoTemporal > rsUsuarioRecaudo!nMaximo Then
                           nMontoTemporal = rsUsuarioRecaudo!nMaximo
                       End If
                       
                       nMonto = nMontoTemporal * ((rsUsuarioRecaudo!nDistCliente) / 100)
                       
                       nnMontoEmp = nMontoTemporal * ((rsUsuarioRecaudo!nDistEmpresa) / 100)
                       
                       txtComision.Text = Format(nMonto, "#,##0.00")
                       nComisionEmpresa = nnMontoEmp
                                                 
                    End If
                    
                ' Caso: Por Concepto->Porcentual===============================================================}
                
            End If
                        
        End If
        
    End If
    
End Sub

Private Sub grdConceptoPagarVC_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)

Dim sColumnas() As String

sColumnas = Split(grdConceptoPagarVC.ColumnasAEditar, "-")

If sColumnas(pnCol) = "X" Then
Cancel = False
MsgBox "No es posible editar este campo", vbInformation, "Aviso"
End If

End Sub

Private Sub grdConceptoPagarVI_GotFocus()
    bFocusGrid = True
End Sub

Private Sub grdConceptoPagarVI_LostFocus()
    bFocusGrid = False
End Sub

Private Sub grdConceptoPagarVI_OnCellChange(pnRow As Long, pnCol As Long)
    
    If grdConceptoPagarVI.Col = 5 Then
    
        If Not IsNumeric(grdConceptoPagarVI.TextMatrix(grdConceptoPagarVI.row, 5)) Then
        
            MsgBox "Valor de la celda no es numerico", vbExclamation, "Aviso"
            grdConceptoPagarVI.TextMatrix(grdConceptoPagarVI.row, 5) = "0.00"
            
        Else
            
            If grdConceptoPagarVI.TextMatrix(grdConceptoPagarVI.row, 5) < 0 Then
            
             MsgBox "Valor de la celda es menor que 0.00", vbExclamation, "Aviso"
             grdConceptoPagarVI.TextMatrix(grdConceptoPagarVI.row, 5) = "0.00"
            
            End If
                        
        End If
    
    End If
    
    calculoSubTotalComisioVI
    
End Sub
Private Sub calculoSubTotalComisioVI()
    Dim nSubTotal As Double
    Dim nI As Integer
    
    For nI = 1 To grdConceptoPagarVI.Rows - 1
        If grdConceptoPagarVI.TextMatrix(nI, 4) = "." Then
            nSubTotal = nSubTotal + CDbl(IIf(Not IsNumeric(grdConceptoPagarVI.TextMatrix(nI, 5)), 0, grdConceptoPagarVI.TextMatrix(nI, 5)))
        End If
    Next
    
    txtSubTotalVI.Text = Format$(nSubTotal, "#,##0.00")
    'calculoComision grdConceptoPagarVI, txtComisionVI, 4, txtSubTotalVI
    calculoComisionNew 'RIRO20150320
    
    ' Suma y muesta el Total
    Dim strValor As Double
    strValor = Math.Round((CDbl(txtSubTotalVI.Text) + CDbl(txtComisionVI.Text)), 2)
    txtTotalVI.Text = Format(strValor, "#,##0.00")
End Sub
 Private Sub grdConceptoPagarVI_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    calculoSubTotalComisioVI
End Sub

Private Sub grdConceptoPagarVI_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    Dim nDeuda As Double 'RIRO20170623
    Dim nImportePago As Double ''RIRO20170623
    
    sColumnas = Split(grdConceptoPagarVI.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "No es posible editar este campo", vbInformation, "Aviso"
    'RIRO20170623 ***
    Else
        If pnCol = 5 Then
            'validando si es numerico
            If Not IsNumeric(grdConceptoPagarVI.TextMatrix(pnRow, pnCol)) Then
                Cancel = False
                MsgBox "El monto ingresado debe ser numérico", vbInformation, "Aviso"
                Exit Sub
            End If
            
            'validando el monto
            nImportePago = CDbl(grdConceptoPagarVI.TextMatrix(pnRow, pnCol))
            nDeuda = CDbl(grdConceptoPagarVI.TextMatrix(pnRow, 6))
            
            If nImportePago > nDeuda Then
                Cancel = False
                MsgBox "El importe de pago no puede ser mayor a la deuda (" & Format(nDeuda, "#0.00") & ")", vbInformation, "Aviso"
                Exit Sub
            End If
            
        End If
    'END RIRO *******
    End If
End Sub

Private Sub grdConceptoPagarVX_GotFocus()
    bFocusGrid = True
End Sub

Private Sub grdConceptoPagarVX_LostFocus()
    bFocusGrid = False
End Sub

Private Sub grdConceptoPagarVX_OnChangeCombo()

    Dim cCodigoImporte As String
    Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
    Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
    Dim nI As Integer
    
    Dim rs As Recordset
    Set rs = New Recordset
        
    On Error GoTo Error
        
    cCodigoImporte = Trim((Right(Trim(grdConceptoPagarVX.TextMatrix(grdConceptoPagarVX.row, 1)), 10)))
    
    Set rs = ClsServicioRecaudo.getMontoConceptoServicio(cCodigoImporte, Trim(txtCodigoBusConvenio))
    
    If Not rs.EOF Then
        grdConceptoPagarVX.TextMatrix(grdConceptoPagarVX.row, 3) = rs!nDeudaActual
        grdConceptoPagarVX.TextMatrix(grdConceptoPagarVX.row, 3) = _
        Format$(grdConceptoPagarVX.TextMatrix(grdConceptoPagarVX.row, 3), "#,##0.00")
        grdConceptoPagarVX.TextMatrix(grdConceptoPagarVX.row, 4) = cCodigoImporte
    Else
        grdConceptoPagarVX.TextMatrix(grdConceptoPagarVX.row, 3) = 0#
    End If

    Dim nSubTotal As Double
    For nI = 1 To grdConceptoPagarVX.Rows - 1
        If grdConceptoPagarVX.TextMatrix(nI, 3) <> "" Then
            nSubTotal = nSubTotal + CDbl(IIf(Not IsNumeric(grdConceptoPagarVX.TextMatrix(nI, 3)), 0, grdConceptoPagarVX.TextMatrix(nI, 3)))
        End If
    Next
    txtSubTotalVX.Text = Format(nSubTotal, "#,##0.00")
    'calculoComision grdConceptoPagarVX, txtComisionVX, -1, txtSubTotalVX
    calculoComisionNew 'RIRO20150320
    
    ' Suma y muesta el Total
    Dim strValor As Double
    strValor = Math.Round((CDbl(txtSubTotalVX.Text) + CDbl(txtComisionVX.Text)), 2)
    txtTotalVX.Text = Format(strValor, "#,##0.00")
      
    Exit Sub
    
Error:
    
    MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub grdConceptoPagarVX_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)

    Dim sColumnas() As String
    sColumnas = Split(grdConceptoPagarVX.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "No es posible editar este campo", vbInformation, "Aviso"
    End If

End Sub
Private Sub txtCodigoBusConvenio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim strCodigo As String
        strCodigo = txtCodigoBusConvenio.Text
        Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
        Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
        Dim rsRecaudo As Recordset
        Set rsRecaudo = New Recordset
        Set rsRecaudo = ClsServicioRecaudo.GetBuscaConvenioXCodigo(strCodigo)
        ' Limpia Detalle de Convevio
        limpiaDetalle
        Set rsUsuarioRecaudo = Nothing
        nMoneda = 1
        nComisionEmpresa = 0#
        If Not rsRecaudo Is Nothing Then
            If Not (rsRecaudo.EOF And rsRecaudo.BOF) Then
                Dim objValidar As COMNCaptaGenerales.NCOMCaptaMovimiento
                Set objValidar = New COMNCaptaGenerales.NCOMCaptaMovimiento
                'Valida el estado de la cuenta, solo permite el pago a cuentas activas
                If Not objValidar.ValidaEstadoCuenta(rsRecaudo!cCtaCod, False) Then
                    MsgBox "La cuenta del convenio seleccionado NO Tiene un estado valido para la operacion", vbExclamation + vbDefaultButton1, "Aviso"
                    limpiaDetalle
                    limipiarCabecera
                    Set ClsServicioRecaudo = Nothing
                    Set objValidar = Nothing
                    Exit Sub
                End If
                'valida que el convenio seleccionado cuente con registros de pago.
                If Mid(rsRecaudo!cCodConvenio, 14, 2) <> "SV" Then
                    'If ClsServicioRecaudo.getBuscarUsuarioRecaudo(, , , rsRecaudo!cCodConvenio).RecordCount = 0 Then
                    If Not ClsServicioRecaudo.getCantidadRegistrosConvenio(rsRecaudo!cCodConvenio) Then
                        If Mid(rsRecaudo!cCodConvenio, 14, 2) = "VP" Then
                            MsgBox "El convenio seleccionado no posee importes a pagar", vbExclamation, "Aviso"
                        Else
                            MsgBox "El convenio seleccionado no posee clientes a pagar", vbExclamation, "Aviso"
                        End If
                        limpiaDetalle
                        limipiarCabecera
                        Exit Sub
                    End If
                End If
                txtCodigoBusConvenio.Text = rsRecaudo!cCodConvenio
                txtNombreConvenio.Text = rsRecaudo!cNombreConvenio
                txtCodigoEmpresa.Text = rsRecaudo!cperscod
                txtNombreEmpresa.Text = rsRecaudo!cPersNombre
                lblDescripcionCodigoVC.Caption = rsRecaudo!cDescripcion
                lblDescripcionCodigoVI.Caption = rsRecaudo!cDescripcion
                lblDescripcionCodigoSV.Caption = rsRecaudo!cDescripcion
                lblDescripcionCodigoVP.Caption = rsRecaudo!cDescripcion
                nMoneda = IIf(Mid(rsRecaudo!cCtaCod, 9, 1) = 1, 1, 2)
                strCuenta = rsRecaudo!cCtaCod
                vistaPestana (txtCodigoBusConvenio.Text)
                If Mid(txtCodigoBusConvenio.Text, 14, 2) = "VP" Then
                    grdConceptoPagarVX.CargaCombo ClsServicioRecaudo.getListaConceptosCobrarXConvenioPV(Trim(txtCodigoBusConvenio.Text))
                End If
                txtCodigoBusConvenio.Locked = True
                cmdGuardar.Enabled = True
                cmdcancelar.Enabled = True
            End If
        Else
             MsgBox "Usted no selecciono ninguna Empresa", vbExclamation, "Aviso"
    '             limipiarCabeceras
        End If
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub
Private Sub txtDOISV_Change()

    If txtDOISV.Text <> "" Then
        
        If Not IsNumeric(txtDOISV.Text) Then
            txtDOISV.Text = ""
            MsgBox "Solo debe ingresar datos numéricos", vbInformation, "Aviso"
            Exit Sub
        End If
        
    End If
    
End Sub

Private Sub txtDOISV_KeyPress(KeyAscii As Integer)
   
    If KeyAscii = 13 Then
        txtOtroCodigoSV.SetFocus
    End If

End Sub

Private Sub txtDOIVX_Change()

    If txtDOIVX.Text <> "" Then
        
        If Not IsNumeric(txtDOIVX.Text) Then
            txtDOIVX.Text = ""
            MsgBox "Solo debe ingresar datos numéricos", vbInformation, "Aviso"
            Exit Sub
        End If
        
    End If

End Sub

Private Sub txtDOIVX_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtOtroCodigoVX.SetFocus
    End If

End Sub

Private Sub txtNombreClienteSV_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cboTipoDOISV.SetFocus
        
    Else
        KeyAscii = Letras(KeyAscii)
    
    End If

End Sub

Private Sub txtNombreClienteVX_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cboTipoDOIVX.SetFocus
        
    Else
        KeyAscii = Letras(KeyAscii)
            
    End If

End Sub

Private Sub txtOtroCodigoSV_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdAgregarSV.SetFocus
    End If

End Sub

Private Sub txtOtroCodigoVX_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = 13 Then
        cmdAgregarVX.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
    
End Sub

Private Function ObtenerImporteMinMax(ByVal pCodConvenio As String, Optional pId As String = "") As ADODB.Recordset
   
    Dim ClsServicioRecaudo As COMDCaptaServicios.DCOMServicioRecaudo
    Dim rsImporteMinMax As ADODB.Recordset
    
    On Error GoTo errImporteMinMax
    
    Set ClsServicioRecaudo = New COMDCaptaServicios.DCOMServicioRecaudo
    Set rsImporteMinMax = ClsServicioRecaudo.getImporteMinMax(pCodConvenio, pId)
    
    Set ObtenerImporteMinMax = rsImporteMinMax
    Set rsImporteMinMax = Nothing
   
    Exit Function
    
errImporteMinMax:
    MsgBox err.Description, vbCritical, "Aviso"
    
End Function



