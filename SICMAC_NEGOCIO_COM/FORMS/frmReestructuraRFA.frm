VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmReestructuraRFA 
   Caption         =   "Restructuración de Creditos RFA"
   ClientHeight    =   8280
   ClientLeft      =   1365
   ClientTop       =   1995
   ClientWidth     =   10875
   Icon            =   "frmReestructuraRFA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   10875
   Begin VB.CommandButton cmdImpresion 
      Caption         =   "Impresion Previa"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5880
      TabIndex        =   52
      Top             =   7755
      Width           =   1440
   End
   Begin VB.CommandButton cmdDeshacer 
      Caption         =   "&Deshacer"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5880
      TabIndex        =   45
      Top             =   7380
      Width           =   1440
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Height          =   375
      Left            =   8025
      TabIndex        =   16
      Top             =   7695
      Width           =   1350
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4425
      TabIndex        =   17
      Top             =   7785
      Width           =   1350
   End
   Begin VB.CommandButton cmdReestructura 
      Caption         =   "&Reestructura"
      Height          =   330
      Left            =   4425
      TabIndex        =   15
      Top             =   7410
      Width           =   1350
   End
   Begin VB.Frame frareest 
      Caption         =   "Reestructurar:"
      Enabled         =   0   'False
      Height          =   810
      Left            =   105
      TabIndex        =   14
      Top             =   7305
      Width           =   4275
      Begin VB.CheckBox chkFechaFija 
         Caption         =   "Fecha Fija"
         Height          =   210
         Left            =   195
         TabIndex        =   44
         Top             =   540
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin MSMask.MaskEdBox txtFechaNew 
         Height          =   330
         Left            =   2850
         TabIndex        =   19
         Top             =   277
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblNroCuo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0000"
         Height          =   300
         Left            =   825
         TabIndex        =   22
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cuota"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   255
         Width           =   420
      End
      Begin VB.Label Label6 
         Caption         =   "Nueva Fecha:"
         Height          =   270
         Left            =   1740
         TabIndex        =   18
         Top             =   307
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdsalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   9390
      TabIndex        =   4
      Top             =   7680
      Width           =   1245
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Datos de Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   10725
      Begin VB.CheckBox CHKRefinanciados 
         Caption         =   "Refinanciados"
         Height          =   255
         Left            =   2910
         TabIndex        =   58
         Top             =   270
         Width           =   1425
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9120
         TabIndex        =   57
         Top             =   210
         Width           =   1395
      End
      Begin SICMACT.TxtBuscar txtBuscaCli 
         Height          =   330
         Left            =   675
         TabIndex        =   56
         Top             =   225
         Width           =   2055
         _extentx        =   3625
         _extenty        =   582
         appearance      =   1
         appearance      =   1
         font            =   "frmReestructuraRFA.frx":030A
         appearance      =   1
         tipobusqueda    =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblnomcli 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   660
         TabIndex        =   1
         Top             =   585
         Width           =   8415
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Ejecutar"
      Height          =   315
      Left            =   4425
      TabIndex        =   23
      Top             =   7425
      Visible         =   0   'False
      Width           =   1350
   End
   Begin TabDlg.SSTab tabRFA 
      Height          =   6180
      Left            =   30
      TabIndex        =   3
      Top             =   990
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   10901
      _Version        =   393216
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
      TabCaption(0)   =   "Credito RFA"
      TabPicture(0)   =   "frmReestructuraRFA.frx":0336
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fgRFA"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Credito CMAC"
      TabPicture(1)   =   "frmReestructuraRFA.frx":0352
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "fgRFC"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Credito DIF"
      TabPicture(2)   =   "frmReestructuraRFA.frx":036E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fgDIF"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).ControlCount=   2
      Begin SICMACT.FlexEdit fgRFA 
         Height          =   4425
         Left            =   90
         TabIndex        =   53
         Top             =   420
         Width           =   10425
         _extentx        =   18389
         _extenty        =   7805
         cols0           =   14
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "Item-Vencimiento-Dias-Nro.Cuota-Capital-Interes-Com.COFIDE-Cuota-Saldo.Cap-Estado-Cap.Pag-Int.Pag-Mora.Pag-Int.Mor."
         encabezadosanchos=   "650-1200-850-1200-1200-1200-1200-1200-1200-900-0-0-0-0"
         font            =   "frmReestructuraRFA.frx":038A
         font            =   "frmReestructuraRFA.frx":03B6
         font            =   "frmReestructuraRFA.frx":03E2
         font            =   "frmReestructuraRFA.frx":040E
         font            =   "frmReestructuraRFA.frx":043A
         fontfixed       =   "frmReestructuraRFA.frx":0466
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-C-C-R-R-R-R-R-C-R-R-R-R"
         formatosedit    =   "0-0-0-0-2-2-2-2-2-0-2-2-2-2"
         textarray0      =   "Item"
         colwidth0       =   645
         rowheight0      =   300
         forecolorfixed  =   8388608
      End
      Begin VB.Frame Frame3 
         Caption         =   "Total Saldos"
         Height          =   1065
         Left            =   -74895
         TabIndex        =   34
         Top             =   4995
         Width           =   10365
         Begin VB.Label lblCapDesDIF 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   315
            Left            =   1035
            TabIndex        =   51
            Top             =   255
            Width           =   1500
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cap.Desemb:"
            Height          =   195
            Left            =   15
            TabIndex        =   50
            Top             =   300
            Width           =   960
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Al reestructurar:"
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
            Left            =   5175
            TabIndex        =   43
            Top             =   45
            Width           =   1380
         End
         Begin VB.Label lblSalIntNewDIF 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   8115
            TabIndex        =   42
            Top             =   570
            Width           =   1500
         End
         Begin VB.Label lblSalCapNewDIf 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   8115
            TabIndex        =   41
            Top             =   195
            Width           =   1500
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Interes : "
            Height          =   195
            Left            =   6990
            TabIndex        =   40
            Top             =   615
            Width           =   1065
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Capital :"
            Height          =   195
            Left            =   6975
            TabIndex        =   39
            Top             =   225
            Width           =   1020
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Capital : "
            Height          =   195
            Left            =   2595
            TabIndex        =   38
            Top             =   315
            Width           =   1065
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Interes : "
            Height          =   195
            Left            =   2580
            TabIndex        =   37
            Top             =   690
            Width           =   1065
         End
         Begin VB.Label lblnSalCapDIF 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   315
            Left            =   3720
            TabIndex        =   36
            Top             =   255
            Width           =   1500
         End
         Begin VB.Label lblSalIntDIF 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   315
            Left            =   3720
            TabIndex        =   35
            Top             =   645
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Total Saldos"
         Height          =   1065
         Left            =   -74910
         TabIndex        =   24
         Top             =   5040
         Width           =   10365
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Cap.Desemb:"
            Height          =   195
            Left            =   135
            TabIndex        =   49
            Top             =   300
            Width           =   960
         End
         Begin VB.Label lblCapDesRFC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   315
            Left            =   1185
            TabIndex        =   48
            Top             =   255
            Width           =   1500
         End
         Begin VB.Label lblSalIntRFC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   315
            Left            =   3930
            TabIndex        =   33
            Top             =   615
            Width           =   1500
         End
         Begin VB.Label lblnSalCapRFC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   315
            Left            =   3930
            TabIndex        =   32
            Top             =   225
            Width           =   1500
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Interes : "
            Height          =   195
            Left            =   2790
            TabIndex        =   31
            Top             =   660
            Width           =   1065
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Capital : "
            Height          =   195
            Left            =   2805
            TabIndex        =   30
            Top             =   285
            Width           =   1065
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Capital :"
            Height          =   195
            Left            =   6975
            TabIndex        =   29
            Top             =   225
            Width           =   1020
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Interes : "
            Height          =   195
            Left            =   6990
            TabIndex        =   28
            Top             =   615
            Width           =   1065
         End
         Begin VB.Label lblSalCapNewRFC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   8115
            TabIndex        =   27
            Top             =   180
            Width           =   1500
         End
         Begin VB.Label lblSalIntNewRFC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   8115
            TabIndex        =   26
            Top             =   570
            Width           =   1500
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Al reestructurar:"
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
            Left            =   5175
            TabIndex        =   25
            Top             =   45
            Width           =   1380
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Total Saldos"
         Height          =   1065
         Left            =   120
         TabIndex        =   5
         Top             =   4935
         Width           =   10365
         Begin VB.Label lblCapDesRFA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   315
            Left            =   1170
            TabIndex        =   47
            Top             =   225
            Width           =   1500
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cap.Desemb:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Al reestructurar:"
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
            Left            =   5175
            TabIndex        =   20
            Top             =   45
            Width           =   1380
         End
         Begin VB.Label lblSalIntNewRFA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   8115
            TabIndex        =   13
            Top             =   570
            Width           =   1500
         End
         Begin VB.Label lblSalCapNewRFA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   8115
            TabIndex        =   12
            Top             =   165
            Width           =   1500
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Interes : "
            Height          =   195
            Left            =   6990
            TabIndex        =   11
            Top             =   615
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Capital :"
            Height          =   195
            Left            =   6975
            TabIndex        =   10
            Top             =   225
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Capital : "
            Height          =   195
            Left            =   2730
            TabIndex        =   9
            Top             =   270
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Interes : "
            Height          =   195
            Left            =   2715
            TabIndex        =   8
            Top             =   660
            Width           =   1065
         End
         Begin VB.Label lblnSalCapRFA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   315
            Left            =   3855
            TabIndex        =   7
            Top             =   225
            Width           =   1500
         End
         Begin VB.Label lblSalIntRFA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   315
            Left            =   3855
            TabIndex        =   6
            Top             =   615
            Width           =   1500
         End
      End
      Begin SICMACT.FlexEdit fgRFC 
         Height          =   4425
         Left            =   -74910
         TabIndex        =   54
         Top             =   480
         Width           =   10425
         _extentx        =   18389
         _extenty        =   7805
         cols0           =   14
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "Item-Vencimiento-Dias-Nro.Cuota-Capital-Interes-Com.COFIDE-Cuota-Saldo.Cap-Estado-Cap.Pag-Int.Pag-Mora.Pag-Int.Mor."
         encabezadosanchos=   "650-1200-850-1200-1200-1200-1200-1200-1200-900-0-0-0-0"
         font            =   "frmReestructuraRFA.frx":0494
         font            =   "frmReestructuraRFA.frx":04C0
         font            =   "frmReestructuraRFA.frx":04EC
         font            =   "frmReestructuraRFA.frx":0518
         font            =   "frmReestructuraRFA.frx":0544
         fontfixed       =   "frmReestructuraRFA.frx":0570
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-C-C-R-R-R-R-R-C-R-R-R-R"
         formatosedit    =   "0-0-0-0-2-2-2-2-2-0-2-2-2-2"
         textarray0      =   "Item"
         colwidth0       =   645
         rowheight0      =   300
         forecolorfixed  =   16711680
      End
      Begin SICMACT.FlexEdit fgDIF 
         Height          =   4470
         Left            =   -74910
         TabIndex        =   55
         Top             =   480
         Width           =   10425
         _extentx        =   18389
         _extenty        =   7885
         cols0           =   14
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "Item-Vencimiento-Dias-Nro.Cuota-Capital-Interes-Com.COFIDE-Cuota-Saldo.Cap-Estado-Cap.Pag-Int.Pag-Mora.Pag-Int.Mor."
         encabezadosanchos=   "650-1200-850-1200-1200-1200-1200-1200-1200-900-0-0-0-0"
         font            =   "frmReestructuraRFA.frx":059E
         font            =   "frmReestructuraRFA.frx":05CA
         font            =   "frmReestructuraRFA.frx":05F6
         font            =   "frmReestructuraRFA.frx":0622
         font            =   "frmReestructuraRFA.frx":064E
         fontfixed       =   "frmReestructuraRFA.frx":067A
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-C-C-R-R-R-R-R-C-R-R-R-R"
         formatosedit    =   "0-0-0-0-2-2-2-2-2-0-2-2-2-2"
         textarray0      =   "Item"
         colwidth0       =   645
         rowheight0      =   300
         forecolorfixed  =   255
      End
   End
End
Attribute VB_Name = "frmReestructuraRFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oRFA  As COMNCredito.NCOMRFA

Dim rsRFA As ADODB.Recordset
Dim rsRFC As ADODB.Recordset
Dim rsDIF As ADODB.Recordset
Dim FilSel As Long
Dim lnTasaRFA As Double
Dim lnPlazoAprRFA As Long
Dim lnTasaRFC As Double
Dim lnPlazoAprRFC As Long
Dim lnTasaDIF As Double
Dim lnPlazoAprDIF As Long
Dim lnCapDesRFA As Currency
Dim lnCapDesRFC As Currency
Dim lnCapDesDIF As Currency
Dim lbReestructurado As Boolean
Dim lsDocID As String
Dim lsDocJur As String

Dim gsCodCtaRFA As String
Dim gsCodCtaRFC As String
Dim gsCodCtaDIF As String
Dim lsImpCopia As String


Dim MatCalendRFC As Variant
Dim MatCalendDIF As Variant
Dim MatCalendRFA As Variant

Dim sCtaCodRFC As String
Dim sCtaCodDIF As String
Dim sCtaCodRFA As String

Private Sub CmdAceptar_Click()
lbReestructurado = False
If ValidaRepro = True Then
    'FilSel = fgRFA.Row
    If FilSel = 0 Then
        MsgBox "Seleccione la cuota a reprogramar", vbInformation, "Aviso"
        Exit Sub
    End If
    If Reprogramación(FilSel) Then
        lbReestructurado = True
        frareest.Enabled = False
        cmdcancelar.Enabled = False
        cmdAceptar.Visible = False
        cmdReestructura.Visible = True
        cmdReestructura.Enabled = False
        cmdDeshacer.Enabled = True
        cmdImpresion.Enabled = True
        cmdGrabar.Enabled = True
    Else
        'CUSCO
        'CargaCalendariosOrig LblCodCli
        If FilSel > 0 Then
            If fgRFA.Rows > 0 Then
                fgRFA.Row = FilSel - 1
                fgRFA.SetFocus
                tabRFA.Tab = 0
            End If
        End If
    End If
Else
    txtFechaNew.SetFocus
    Exit Sub
End If
End Sub

Private Sub cmdCancelar_Click()
    Me.frareest.Enabled = False
    Me.cmdAceptar.Visible = False
    Me.cmdReestructura.Visible = True
    Me.cmdcancelar.Enabled = False
    Me.cmdGrabar.Enabled = False
    Me.cmdDeshacer.Enabled = False
    cmdImpresion.Enabled = False
End Sub

Private Sub cmdDeshacer_Click()
CargaCalendariosOrig txtBuscaCli.Text, True
If FilSel > 0 And FilSel <= fgRFA.Rows Then
    fgRFA.Row = FilSel
End If
fgRFA.SetFocus
tabRFA.Tab = 0
cmdDeshacer.Enabled = False
Me.cmdReestructura.Enabled = True
lbReestructurado = False
Me.cmdGrabar.Enabled = False
Me.cmdImpresion.Enabled = False
lblSalCapNewDIf = "0.00"
lblSalCapNewRFA = "0.00"
lblSalCapNewRFC = "0.00"
lblSalIntNewDIF = "0.00"
lblSalIntNewRFA = "0.00"
lblSalIntNewRFC = "0.00"

End Sub

Private Sub cmdGrabar_Click()
    GrabaReestructuracion
End Sub
Sub GrabaReestructuracion()
'Dim sql As String
'Dim i As Integer
'Dim lbTran As Boolean
Dim oNCredRFA As COMNCredito.NCOMRFA
Dim oDCredRFA As COMDCredito.DCOMRFA
Dim nNroCalen As Integer
'On Error GoTo ErrorGrabar
If MsgBox("Desea grabar la reestructuración del calendario ??? ", vbYesNo + vbQuestion, "Aviso") = vbNo Then
    Exit Sub
End If


'dbBase.BeginTrans
'lbTran = True
''guardamos el calendario original antes de reestructuración
'If Not rsRFA Is Nothing Then
'    rsRFA.MoveFirst
'    sql = "DELETE FROM " & gsRutaPrinc & gsBaseCred & "KPYDPPGRESTRFA WHERE cCodCta ='" & rsRFA!cCodCta & "'"
'    dbBase.Execute sql
'    Do While Not rsRFA.EOF
'        sql = "INSERT INTO " & gsRutaPrinc & gsBaseCred & "KPYDPPGRESTRFA " _
'        & " (cCodCta, dFecVen, cEstado, cTipOpe, cNroCuo, nCapita, " _
'        & " nIntere, nCapPag, nIntPag, nIntMor, nMorPag, nCofide, nCofPag, " _
'        & " cCodUsu, dFecMod, cFlag, dFecPag, nIntDif, nIntdPag, nOtrPag ) " _
'        & " VALUES('" & rsRFA!cCodCta & "',CTOD('" & Format(rsRFA!dFecVen, "mm/dd/yyyy") & "'),'" _
'        & rsRFA!cEstado & "','" & rsRFA!cTipOpe & "','" & rsRFA!cNrocuo & "'," & rsRFA!nCapita & "," _
'        & rsRFA!nIntere & "," & rsRFA!nCapPag & "," & rsRFA!nIntPag & "," & rsRFA!nIntMor & "," _
'        & rsRFA!nMorPag & "," & rsRFA!nCofide & "," & rsRFA!nCofPag & ",'" & gsCodUser & "'," _
'        & "ctod('" & Format(gdFecSis, "mm/dd/yyyy") & "'),'R',ctod('" & Format(gdFecSis, "mm/dd/yyyy") & "'),0,0,0)"
'
'        dbBase.Execute sql
'        rsRFA.MoveNext
'    Loop
'End If
'
'If Not rsRFC Is Nothing Then
'    rsRFC.MoveFirst
'    sql = "DELETE FROM " & gsRutaPrinc & gsBaseCred & "KPYDPPGRESTRFA WHERE cCodCta ='" & rsRFC!cCodCta & "'"
'    dbBase.Execute sql
'    Do While Not rsRFC.EOF
'        sql = "INSERT INTO " & gsRutaPrinc & gsBaseCred & "KPYDPPGRESTRFA " _
'        & " (cCodCta, dFecVen, cEstado, cTipOpe, cNroCuo, nCapita, " _
'        & " nIntere, nCapPag, nIntPag, nIntMor, nMorPag, nCofide, nCofPag, " _
'        & " cCodUsu, dFecMod, cFlag, dFecPag, nIntDif, nIntdPag, nOtrPag ) " _
'        & " VALUES('" & rsRFC!cCodCta & "',CTOD('" & Format(rsRFC!dFecVen, "mm/dd/yyyy") & "'),'" _
'        & rsRFC!cEstado & "','" & rsRFC!cTipOpe & "','" & rsRFC!cNrocuo & "'," & rsRFC!nCapita & "," _
'        & rsRFC!nIntere & "," & rsRFC!nCapPag & "," & rsRFC!nIntPag & "," & rsRFC!nIntMor & "," _
'        & rsRFC!nMorPag & "," & rsRFC!nCofide & "," & rsRFC!nCofPag & ",'" & gsCodUser & "'," _
'        & "ctod('" & Format(gdFecSis, "mm/dd/yyyy") & "'),'R',ctod('" & Format(gdFecSis, "mm/dd/yyyy") & "'),0,0,0)"
'
'        dbBase.Execute sql
'        rsRFC.MoveNext
'    Loop
'End If
'
'If Not rsDIF Is Nothing Then
'    rsDIF.MoveFirst
'    sql = "DELETE FROM " & gsRutaPrinc & gsBaseCred & "KPYDPPGRESTRFA WHERE cCodCta ='" & rsDIF!cCodCta & "'"
'    dbBase.Execute sql
'    Do While Not rsDIF.EOF
'        sql = "INSERT INTO " & gsRutaPrinc & gsBaseCred & "KPYDPPGRESTRFA " _
'        & " (cCodCta, dFecVen, cEstado, cTipOpe, cNroCuo, nCapita, " _
'        & " nIntere, nCapPag, nIntPag, nIntMor, nMorPag, nCofide, nCofPag, " _
'        & " cCodUsu, dFecMod, cFlag, dFecPag, nIntDif, nIntdPag, nOtrPag ) " _
'        & " VALUES('" & rsDIF!cCodCta & "',CTOD('" & Format(rsDIF!dFecVen, "mm/dd/yyyy") & "'),'" _
'        & rsDIF!cEstado & "','" & rsDIF!cTipOpe & "','" & rsDIF!cNrocuo & "'," & rsDIF!nCapita & "," _
'        & rsDIF!nIntere & "," & rsDIF!nCapPag & "," & rsDIF!nIntPag & "," & rsDIF!nIntMor & "," _
'        & rsDIF!nMorPag & "," & rsDIF!nCofide & "," & rsDIF!nCofPag & ",'" & gsCodUser & "'," _
'        & "ctod('" & Format(gdFecSis, "mm/dd/yyyy") & "'),'R',ctod('" & Format(gdFecSis, "mm/dd/yyyy") & "'),0,0,0)"
'
'        dbBase.Execute sql
'        rsDIF.MoveNext
'    Loop
'End If
'
''actualizamos el plan de pagos reestructurado
''actualizamos plan de pagos RFA
'For i = 1 To fgRFA.Rows - 1
'    If fgRFA.TextMatrix(fgRFA.Row, 9) <> "P" Then
'        sql = "UPDATE " & gsRutaPrinc & gsBaseCred & "KPYDPPG " _
'        & " SET dFecVen = CTOD('" & Format(fgRFA.TextMatrix(i, 1), "mm/dd/yyyy") & "')," _
'        & "     nCapita = " & CCur(fgRFA.TextMatrix(i, 4)) & "," _
'        & "     nIntere = " & CCur(fgRFA.TextMatrix(i, 5)) & "," _
'        & "     cCodusu = '" & gsCodUser & "'," _
'        & "     dFecMod = CTOD('" & Format(gdFecSis, "mm/dd/yyyy") & "') " _
'        & " WHERE cCodCta ='" & gsCodCtaRFA & "' AND cNroCuo ='" & fgRFA.TextMatrix(i, 3) & "' and cTipOpe ='N' "
'
'        dbBase.Execute sql
'    End If
'Next
'
''actualizamos plan de pagos RFC
'For i = 1 To fgRFC.Rows - 1
'    If fgRFC.TextMatrix(fgRFC.Row, 9) <> "P" Then
'        sql = "UPDATE " & gsRutaPrinc & gsBaseCred & "KPYDPPG " _
'        & " SET dFecVen = CTOD('" & Format(fgRFC.TextMatrix(i, 1), "mm/dd/yyyy") & "')," _
'        & "     nCapita = " & CCur(fgRFC.TextMatrix(i, 4)) & "," _
'        & "     nIntere = " & CCur(fgRFC.TextMatrix(i, 5)) & "," _
'        & "     cCodusu = '" & gsCodUser & "'," _
'        & "     dFecMod = CTOD('" & Format(gdFecSis, "mm/dd/yyyy") & "') " _
'        & " WHERE cCodCta ='" & gsCodCtaRFC & "' AND cNroCuo ='" & fgRFC.TextMatrix(i, 3) & "' and cTipOpe ='N' "
'
'        dbBase.Execute sql
'    End If
'Next
'
''actualizamos plan de pagos DIF
'For i = 1 To fgDIF.Rows - 1
'    If fgDIF.TextMatrix(fgDIF.Row, 9) <> "P" Then
'        sql = "UPDATE " & gsRutaPrinc & gsBaseCred & "KPYDPPG " _
'        & " SET dFecVen = CTOD('" & Format(fgDIF.TextMatrix(i, 1), "mm/dd/yyyy") & "')," _
'        & "     nCapita = " & CCur(fgDIF.TextMatrix(i, 4)) & "," _
'        & "     nIntere = " & CCur(fgDIF.TextMatrix(i, 5)) & "," _
'        & "     cCodusu = '" & gsCodUser & "'," _
'        & "     dFecMod = CTOD('" & Format(gdFecSis, "mm/dd/yyyy") & "') " _
'        & " WHERE cCodCta ='" & gsCodCtaDIF & "' AND cNroCuo ='" & fgDIF.TextMatrix(i, 3) & "' and cTipOpe ='N' "
'
'        dbBase.Execute sql
'    End If
'Next
''actualizamos el maestro de creditos
'sql = "UPDATE " & gsRutaPrinc & gsBaseCred & "KPYMCRE " _
'    & " SET nIntApr = nIntPag  + " & CCur(lblSalIntNewRFA) & "," _
'    & "     cCodusu = '" & gsCodUser & "'," _
'    & "     dFecMod = CTOD('" & Format(gdFecSis, "mm/dd/yyyy") & "') " _
'    & " WHERE cCodCta ='" & gsCodCtaRFA & "' "
'dbBase.Execute sql
'
'sql = "UPDATE " & gsRutaPrinc & gsBaseCred & "KPYMCRE " _
'    & " SET nIntApr = nIntPag  + " & CCur(lblSalIntNewRFC) & "," _
'    & "     cCodusu = '" & gsCodUser & "'," _
'    & "     dFecMod = CTOD('" & Format(gdFecSis, "mm/dd/yyyy") & "') " _
'    & " WHERE cCodCta ='" & gsCodCtaRFC & "' "
'dbBase.Execute sql
'
'sql = "UPDATE " & gsRutaPrinc & gsBaseCred & "KPYMCRE " _
'    & " SET nIntApr = nIntPag  + " & CCur(lblSalIntNewDIF) & "," _
'    & "     cCodusu = '" & gsCodUser & "'," _
'    & "     dFecMod = CTOD('" & Format(gdFecSis, "mm/dd/yyyy") & "') " _
'    & " WHERE cCodCta ='" & gsCodCtaDIF & "' "
'dbBase.Execute sql
'
'dbBase.CommitTrans
'lbTran = False
'
'Dim lnYes As Long
'lnYes = vbYes
'Do While lnYes <> vbNo
'    lsImpCopia = ""
'    PrtCal1
'    PrtCal2
'    lnYes = MsgBox("Desea reimprimir documentos de Reestructuración RFA ???", vbYesNo + vbQuestion, "AVISO")
'Loop

Set oDCredRFA = New COMDCredito.DCOMRFA
nNroCalen = oDCredRFA.GetNroCalend(sCtaCodRFC)
Set oDCredRFA = Nothing
nNroCalen = nNroCalen + 1
Call DistribuyeMatrices

Set oNCredRFA = New COMNCredito.NCOMRFA
Call oNCredRFA.ReprogramarCalendario(MatCalendDIF, MatCalendRFC, MatCalendRFA, nNroCalen, sCtaCodDIF, sCtaCodRFC, _
                                       sCtaCodRFA)
    
Set oNCredRFA = Nothing

Dim lnYes As Long
lnYes = vbYes
Do While lnYes <> vbNo
    lsImpCopia = ""
    PrtCal1
    PrtCal2
    lnYes = MsgBox("Desea reimprimir documentos de Reestructuración RFA ???", vbYesNo + vbQuestion, "AVISO")
Loop
    
Call cmdDeshacer_Click
'Me.LblCodCli = ""
Me.lblnomcli = ""
Me.fgDIF.Clear
Me.fgDIF.Rows = 2
Me.fgDIF.FormaCabecera

Me.fgRFA.Clear
Me.fgRFA.Rows = 2
Me.fgRFA.FormaCabecera

Me.fgRFC.Clear
Me.fgRFC.Rows = 2
Me.fgRFC.FormaCabecera

Me.lblCapDesDIF = "0.00"
Me.lblCapDesRFA = "0.00"
Me.lblCapDesRFC = "0.00"
Me.lblNroCuo = ""
Me.lblnSalCapDIF = "0.00"
Me.lblnSalCapRFA = "0.00"
Me.lblnSalCapRFC = "0.00"
Me.lblSalCapNewDIf = "0.00"
Me.lblSalCapNewRFA = "0.00"
Me.lblSalCapNewRFC = "0.00"
Me.lblSalIntDIF = "0.00"
Me.lblSalIntNewDIF = "0.00"
Me.lblSalIntNewRFA = "0.00"
Me.lblSalIntNewRFC = "0.00"
Me.lblSalIntRFA = "0.00"
Me.lblSalIntRFC = "0.00"


'generamos impresion del nuevo plan de pagos.
Exit Sub
ErrorGrabar:
    'If lbTran Then
    '    dbBase.RollbackTrans
    'End If
    MsgBox Err.Description & " [" & Err.Number & "]", vbInformation, "AVISO"
End Sub

Private Sub cmdReestructura_Click()
If fgRFA.TextMatrix(0, 0) = "" Then Exit Sub
If tabRFA.Tab <> 0 Then
    MsgBox "Seleccione una cuota desde el calendario RFA como partida", vbInformation, "Aviso"
    tabRFA.Tab = 0
    fgRFA.SetFocus
    Exit Sub
End If
If fgRFA.TextMatrix(fgRFA.Row, 9) = "P" Then
    MsgBox "Cuota se encuentra cancelada, no podrá continuar", vbInformation, "Aviso"
    Exit Sub
End If
cmdReestructura.Visible = False
cmdAceptar.Visible = True
cmdcancelar.Enabled = True
frareest.Enabled = True
txtFechaNew.SetFocus
lblNroCuo = fgRFA.TextMatrix(fgRFA.Row, 3)
FilSel = fgRFA.Row
If FilSel = 0 Then
    MsgBox "Por favor seleccione la cuota a reestructurar", vbInformation, "Aviso"
End If

End Sub

Private Function ValidaRepro() As Boolean
Dim Cad As String
    'Valida Exista Elementos
    If Me.fgRFA.TextMatrix(0, 0) = "" Then
        MsgBox "No Existen Elementos en la Lista", vbInformation, "Aviso"
        ValidaRepro = False
        Exit Function
    End If
    'Valida Fecha de Reprogramacion
    Cad = ValidaFecha(txtFechaNew)
    If Len(Trim(Cad)) > 0 Then
        MsgBox Cad, vbInformation, "Aviso"
        ValidaRepro = False
        Exit Function
    End If
    'Valida que Cuota No este Cancelada
    If fgRFA.TextMatrix(FilSel, 9) = "P" Then
        MsgBox "No Se Puede Reprogramar la Cuota seleccionada del RFA que Ya Esta Cancelada", vbInformation, "Aviso"
        ValidaRepro = False
        Exit Function
    End If
    'VALIDAMOS SI ALGUNA DE LAS CUOTAS DEL RFA/RFC/DIF ESTA CANCELADA
    
    'Valida que Cuota RFC No este Cancelada
    If fgRFC.TextMatrix(FilSel, 9) = "1" Then
        MsgBox "No Se Puede Reprogramar la Cuota seleccionada del RFC Ya Esta Cancelada", vbInformation, "Aviso"
        ValidaRepro = False
        Exit Function
    End If
    'Valida que Cuota DIF No este Cancelada
    If fgDIF.TextMatrix(FilSel, 9) = "1" Then
        MsgBox "No Se Puede Reprogramar la Cuota seleccionada del DIF Ya Esta Cancelada", vbInformation, "Aviso"
        ValidaRepro = False
        Exit Function
    End If
    'Valida que No se Reprograme antes de la fecha de Vencimiento
    If FilSel - 1 >= 1 Then
        If CDate(txtFechaNew) <= CDate(fgRFA.TextMatrix(FilSel - 1, 1)) Then
            MsgBox "No Se Puede Reprogramar la Fecha de Vencimiento de una Cuota " & _
                   " A Una Fecha Anterior a Esta", vbInformation, "Aviso"
            ValidaRepro = False
            Exit Function
        End If
    Else
        ValidaRepro = False
        Exit Function
    End If
    If (CDate(fgRFA.TextMatrix(FilSel, 1)) <> CDate(fgRFC.TextMatrix(FilSel, 1))) Or _
        (CDate(fgRFA.TextMatrix(FilSel, 1)) <> CDate(fgDIF.TextMatrix(FilSel, 1))) Or _
        (CDate(fgRFC.TextMatrix(FilSel, 1)) <> CDate(fgDIF.TextMatrix(FilSel, 1))) Then
        MsgBox "Fecha de Vencimiento de la cuota seleccionada es diferente en alguno de los tres creditos [RFA,RFC,DIF]" & Chr(13) & "Por favor Revice", vbInformation, "Aviso"
        ValidaRepro = False
        Exit Function
    End If
    If CDate(txtFechaNew) < gdFecSis Then
        MsgBox "Fecha de reestructuración no puede ser menor a la fecha del sistemas", vbInformation, "aviso"
        ValidaRepro = False
        Exit Function
    End If
    
    ValidaRepro = True
End Function

Function ReprogramaRFA(ByVal FilSel As Long) As Boolean

Dim NuevaFecha As Date
Dim TasaReal As Double
Dim I As Long
Dim ContPlan As Long
Dim lnDiasTrans As Long
Dim lnSalCap As Currency
Dim lnAnioAnt As Long
Dim ldCuota As Date
Dim NuevoInteres  As Double
Dim lnSalCapNew As Currency
Dim lnSalIntNew As Currency
Dim TotalCap As Double
Dim TotalInt As Double
Dim ldFecAnt As Date
'************* REPROGRAMAMOS CREDITO RFA *******************************************************

ReprogramaRFA = False

ContPlan = fgRFA.Rows
TotalInt = 0
TotalCap = 0

ReprogramaRFA = False

NuevaFecha = CDate(txtFechaNew)

lnDiasTrans = DateDiff("d", fgRFA.TextMatrix(FilSel - 1, 1), NuevaFecha)
TasaReal = InteresReal(lnTasaRFA / 100, lnDiasTrans)
If FilSel = 1 Then
    NuevoInteres = TasaReal * lnCapDesRFA
Else
    NuevoInteres = TasaReal * CCur(fgRFA.TextMatrix(FilSel - 1, 8))
End If
NuevoInteres = Format(NuevoInteres, "#0.00")

If NuevoInteres > CCur(fgRFA.TextMatrix(FilSel, 7)) Then
    MsgBox "Calculo de Nuevo interés en credito RFA :[" & Format(NuevoInteres, "#,#0.00") & "] sobrepasa el monto de la cuota asignada= [" & Format(fgRFA.TextMatrix(FilSel, 7), "#,#0.00") & "] " & Chr(13) & "Por favor modifique la nueva fecha a una menor que la ingresada", vbInformation, "Aviso"
    txtFechaNew.SetFocus
    Exit Function
End If
fgRFA.TextMatrix(FilSel, 1) = NuevaFecha
fgRFA.TextMatrix(FilSel, 2) = lnDiasTrans
fgRFA.TextMatrix(FilSel, 4) = Format(CCur(fgRFA.TextMatrix(FilSel, 7)) - NuevoInteres, "#,#0.00") 'NUEVO CAPITAL
fgRFA.TextMatrix(FilSel, 5) = Format(NuevoInteres, "#,#0.00") 'NUEVO INTERES
fgRFA.TextMatrix(FilSel, 7) = Format(CCur(fgRFA.TextMatrix(FilSel, 4)) + CCur(fgRFA.TextMatrix(FilSel, 5)), "#,#0.00")  ' NUEVA CUOTA
fgRFA.TextMatrix(FilSel, 8) = Format(CCur(fgRFA.TextMatrix(FilSel - 1, 8)) - CCur(fgRFA.TextMatrix(FilSel, 4)), "#,#0.00")
lnSalCap = fgRFA.TextMatrix(FilSel, 8)

fgRFA.TextMatrix(FilSel, 13) = 0
fgRFA.TextMatrix(FilSel, 12) = 0
fgRFA.ForeColorRow vbBlue

ldFecAnt = fgRFA.TextMatrix(FilSel, 1)
lnAnioAnt = Year(ldFecAnt)
ldCuota = ldFecAnt
For I = FilSel + 1 To ContPlan - 2
    If chkFechaFija.value = 0 Then  'PERIODO FIJO
        fgRFA.TextMatrix(I, 1) = DateAdd("d", lnPlazoAprRFA, ldFecAnt)
    Else
        ldCuota = CDate(Format(Day(NuevaFecha), "00") + "/" + Format(Month(NuevaFecha), "00") + "/" + Format(lnAnioAnt + 1, "0000"))
        fgRFA.TextMatrix(I, 1) = ldCuota
    End If
    lnDiasTrans = DateDiff("D", ldFecAnt, fgRFA.TextMatrix(I, 1))
    TasaReal = InteresReal(lnTasaRFA / 100, lnDiasTrans)
    NuevoInteres = TasaReal * lnSalCap
    
    fgRFA.TextMatrix(I, 2) = lnDiasTrans
    fgRFA.TextMatrix(I, 4) = Format(CCur(fgRFA.TextMatrix(I, 7)) - NuevoInteres, "#,#0.00") 'NUEVO CAPITAL
    fgRFA.TextMatrix(I, 5) = Format(NuevoInteres, "#,#0.00") 'NUEVO INTERES
    fgRFA.TextMatrix(I, 7) = Format(CCur(fgRFA.TextMatrix(I, 4)) + CCur(fgRFA.TextMatrix(I, 5)), "#,#0.00")
    fgRFA.TextMatrix(I, 8) = Format(CCur(fgRFA.TextMatrix(I - 1, 8)) - CCur(fgRFA.TextMatrix(I, 4)), "#,#0.00")
    fgRFA.Row = I
    fgRFA.ForeColorRow vbBlue
    
    lnSalCap = fgRFA.TextMatrix(I, 8)

    fgRFA.TextMatrix(I, 13) = 0
    fgRFA.TextMatrix(I, 12) = 0
        
    ldFecAnt = fgRFA.TextMatrix(I, 1)
    lnAnioAnt = Year(ldFecAnt)
Next I
'ahora calculamos el capital para la cuota final
lnDiasTrans = DateDiff("D", ldFecAnt, fgRFA.TextMatrix(ContPlan - 1, 1))
TasaReal = InteresReal(lnTasaRFA / 100, lnDiasTrans)
NuevoInteres = TasaReal * lnSalCap
fgRFA.TextMatrix(ContPlan - 1, 2) = lnDiasTrans
fgRFA.TextMatrix(ContPlan - 1, 4) = Format(lnSalCap, "#,#0.00") 'NUEVO CAPITAL
fgRFA.TextMatrix(ContPlan - 1, 5) = Format(NuevoInteres, "#,#0.00") 'NUEVO INTERES
fgRFA.TextMatrix(ContPlan - 1, 7) = Format(CCur(fgRFA.TextMatrix(ContPlan - 1, 4)) + CCur(fgRFA.TextMatrix(ContPlan - 1, 5)), "#,#0.00")
fgRFA.TextMatrix(ContPlan - 1, 8) = Format(CCur(fgRFA.TextMatrix(ContPlan - 2, 8)) - CCur(fgRFA.TextMatrix(ContPlan - 1, 4)), "#,#0.00")

fgRFA.Row = ContPlan - 1
fgRFA.ForeColorRow vbBlue
lnSalCap = fgRFA.TextMatrix(ContPlan - 1, 8)

lnSalCapNew = 0
lnSalIntNew = 0
For I = 1 To ContPlan - 1
    If fgRFA.TextMatrix(I, 9) <> "P" Then
        lnSalCapNew = lnSalCapNew + CCur(fgRFA.TextMatrix(I, 4))
        lnSalIntNew = lnSalIntNew + CCur(fgRFA.TextMatrix(I, 5))
    End If
Next

lblSalCapNewRFA = Format(lnSalCapNew, "#,#0.00")
lblSalIntNewRFA = Format(lnSalIntNew, "#,#0.00")


ReprogramaRFA = True
End Function

Function ReprogramaRFC(ByVal FilSel As Long) As Boolean

Dim NuevaFecha As Date
Dim TasaReal As Double
Dim I As Long
Dim ContPlan As Long
Dim lnDiasTrans As Long
Dim lnSalCap As Currency
Dim lnAnioAnt As Long
Dim ldCuota As Date
Dim NuevoInteres  As Double
Dim lnSalCapNew As Currency
Dim lnSalIntNew As Currency
Dim TotalCap As Double
Dim TotalInt As Double
Dim ldFecAnt As Date
'**************************************************************************************************
'****************************  REPROGRAMAMOS EL CREDITO RFC    *********************************

ReprogramaRFC = False

ContPlan = fgRFC.Rows
TotalInt = 0
TotalCap = 0
lnSalCapNew = 0
lnSalIntNew = 0

NuevaFecha = CDate(txtFechaNew)
lnDiasTrans = DateDiff("d", fgRFC.TextMatrix(FilSel - 1, 1), NuevaFecha)
TasaReal = InteresReal(lnTasaRFC / 100, lnDiasTrans)
If FilSel = 1 Then
    NuevoInteres = TasaReal * lnCapDesRFC
Else
    NuevoInteres = TasaReal * CCur(fgRFC.TextMatrix(FilSel - 1, 8))
End If
NuevoInteres = Format(NuevoInteres, "#0.00")

If NuevoInteres > CCur(fgRFC.TextMatrix(FilSel, 7)) Then
    MsgBox "Calculo de Nuevo interés en crédito RFC :[" & Format(NuevoInteres, "#,#0.00") & "] sobrepasa el monto de la cuota asignada= [" & Format(fgRFC.TextMatrix(FilSel, 7), "#,#0.00") & "] " & Chr(13) & "Por favor modifique la nueva fecha a una menor que la ingresada", vbInformation, "Aviso"
    txtFechaNew.SetFocus
    Exit Function
End If


fgRFC.TextMatrix(FilSel, 1) = NuevaFecha
fgRFC.TextMatrix(FilSel, 2) = lnDiasTrans
fgRFC.TextMatrix(FilSel, 4) = Format(CCur(fgRFC.TextMatrix(FilSel, 7)) - NuevoInteres, "#,#0.00") 'NUEVO CAPITAL
fgRFC.TextMatrix(FilSel, 5) = Format(NuevoInteres, "#,#0.00") 'NUEVO INTERES
fgRFC.TextMatrix(FilSel, 7) = Format(CCur(fgRFC.TextMatrix(FilSel, 4)) + CCur(fgRFC.TextMatrix(FilSel, 5)) + CCur(fgRFC.TextMatrix(FilSel, 6)), "#,#0.00") ' NUEVA CUOTA
fgRFC.TextMatrix(FilSel, 8) = Format(CCur(fgRFC.TextMatrix(FilSel - 1, 8)) - CCur(fgRFC.TextMatrix(FilSel, 4)), "#,#0.00")

fgRFC.Row = FilSel
fgRFC.ForeColorRow vbBlue
lnSalCap = fgRFC.TextMatrix(FilSel, 8)

fgRFC.TextMatrix(FilSel, 13) = 0
fgRFC.TextMatrix(FilSel, 12) = 0
ldFecAnt = fgRFC.TextMatrix(FilSel, 1)
lnAnioAnt = Year(ldFecAnt)
ldCuota = ldFecAnt
For I = FilSel + 1 To ContPlan - 2
    If chkFechaFija.value = 0 Then  'PERIODO FIJO
        fgRFC.TextMatrix(I, 1) = DateAdd("d", lnPlazoAprRFA, ldFecAnt)
    Else
        ldCuota = CDate(Format(Day(NuevaFecha), "00") + "/" + Format(Month(NuevaFecha), "00") + "/" + Format(lnAnioAnt + 1, "0000"))
        fgRFC.TextMatrix(I, 1) = ldCuota
    End If
    lnDiasTrans = DateDiff("D", ldFecAnt, fgRFC.TextMatrix(I, 1))
    TasaReal = InteresReal(lnTasaRFC / 100, lnDiasTrans)
    NuevoInteres = TasaReal * lnSalCap
    
    fgRFC.TextMatrix(I, 2) = lnDiasTrans
    fgRFC.TextMatrix(I, 4) = Format(CCur(fgRFC.TextMatrix(I, 7)) - NuevoInteres, "#,#0.00") 'NUEVO CAPITAL
    fgRFC.TextMatrix(I, 5) = Format(NuevoInteres, "#,#0.00") 'NUEVO INTERES
    fgRFC.TextMatrix(I, 7) = Format(CCur(fgRFC.TextMatrix(I, 4)) + CCur(fgRFC.TextMatrix(I, 5)) + CCur(fgRFC.TextMatrix(I, 6)), "#,#0.00")
    fgRFC.TextMatrix(I, 8) = Format(CCur(fgRFC.TextMatrix(I - 1, 8)) - CCur(fgRFC.TextMatrix(I, 4)), "#,#0.00")
    
    fgRFC.Row = I
    fgRFC.ForeColorRow vbBlue
    lnSalCap = fgRFC.TextMatrix(I, 8)

    fgRFC.TextMatrix(I, 13) = 0
    fgRFC.TextMatrix(I, 12) = 0
    
    ldFecAnt = fgRFC.TextMatrix(I, 1)
    lnAnioAnt = Year(ldFecAnt)
Next I
'ahora calculamos el capital para la cuota final
lnDiasTrans = DateDiff("D", ldFecAnt, fgRFC.TextMatrix(ContPlan - 1, 1))
TasaReal = InteresReal(lnTasaRFC / 100, lnDiasTrans)
NuevoInteres = TasaReal * lnSalCap
fgRFC.TextMatrix(ContPlan - 1, 2) = lnDiasTrans
fgRFC.TextMatrix(ContPlan - 1, 4) = Format(lnSalCap, "#,#0.00") 'NUEVO CAPITAL
fgRFC.TextMatrix(ContPlan - 1, 5) = Format(NuevoInteres, "#,#0.00") 'NUEVO INTERES
fgRFC.TextMatrix(ContPlan - 1, 7) = Format(CCur(fgRFC.TextMatrix(ContPlan - 1, 4)) + CCur(fgRFC.TextMatrix(ContPlan - 1, 5)) + CCur(fgRFC.TextMatrix(ContPlan - 1, 6)), "#,#0.00")
fgRFC.TextMatrix(ContPlan - 1, 8) = Format(CCur(fgRFC.TextMatrix(ContPlan - 2, 8)) - CCur(fgRFC.TextMatrix(ContPlan - 1, 4)), "#,#0.00")

lnSalCapNew = lnSalCapNew + CCur(fgRFC.TextMatrix(ContPlan - 1, 4))
lnSalIntNew = lnSalIntNew + CCur(fgRFC.TextMatrix(ContPlan - 1, 5))

fgRFC.Row = ContPlan - 1
fgRFC.ForeColorRow vbBlue
lnSalCap = fgRFC.TextMatrix(ContPlan - 1, 8)

lnSalCapNew = 0
lnSalIntNew = 0
For I = 1 To ContPlan - 1
    If fgRFC.TextMatrix(I, 9) <> "P" Then
        lnSalCapNew = lnSalCapNew + CCur(fgRFC.TextMatrix(I, 4))
        lnSalIntNew = lnSalIntNew + CCur(fgRFC.TextMatrix(I, 5))
    End If
Next

lblSalCapNewRFC = Format(lnSalCapNew, "#,#0.00")
lblSalIntNewRFC = Format(lnSalIntNew, "#,#0.00")

ReprogramaRFC = True
End Function

Function ReprogramaDIF(ByVal FilSel As Long) As Boolean

Dim NuevaFecha As Date
Dim TasaReal As Double
Dim I As Long
Dim ContPlan As Long
Dim lnDiasTrans As Long
Dim lnSalCap As Currency
Dim lnAnioAnt As Long
Dim ldCuota As Date
Dim NuevoInteres  As Double
Dim lnSalCapNew As Currency
Dim lnSalIntNew As Currency
Dim TotalCap As Double
Dim TotalInt As Double
Dim ldFecAnt As Date

'**************************************************************************************************
'****************************  REPROGRAMAMOS EL CREDITO DIF    *********************************
'*****************************************************************************************************

ReprogramaDIF = False

ContPlan = fgDIF.Rows
TotalInt = 0
TotalCap = 0
lnSalCapNew = 0
lnSalIntNew = 0

NuevaFecha = CDate(txtFechaNew)
lnDiasTrans = DateDiff("d", fgRFC.TextMatrix(FilSel - 1, 1), NuevaFecha)
TasaReal = InteresReal(lnTasaDIF / 100, lnDiasTrans)
If FilSel = 1 Then
    NuevoInteres = TasaReal * lnCapDesDIF
Else
    NuevoInteres = TasaReal * CCur(fgDIF.TextMatrix(FilSel - 1, 8))
End If
NuevoInteres = Format(NuevoInteres, "#0.00")

If NuevoInteres > CCur(fgDIF.TextMatrix(FilSel, 7)) Then
    MsgBox "Calculo de Nuevo interés en crédito DIF :[" & Format(NuevoInteres, "#,#0.00") & "] sobrepasa el monto de la cuota asignada= [" & Format(fgDIF.TextMatrix(FilSel, 7), "#,#0.00") & "] " & Chr(13) & "Por favor modifique la nueva fecha a una menor que la ingresada", vbInformation, "Aviso"
    txtFechaNew.SetFocus
    Exit Function
End If

fgDIF.TextMatrix(FilSel, 1) = NuevaFecha
fgDIF.TextMatrix(FilSel, 2) = lnDiasTrans
fgDIF.TextMatrix(FilSel, 4) = Format(CCur(fgDIF.TextMatrix(FilSel, 7)) - NuevoInteres, "#,#0.00") 'NUEVO CAPITAL
fgDIF.TextMatrix(FilSel, 5) = Format(NuevoInteres, "#,#0.00") 'NUEVO INTERES
fgDIF.TextMatrix(FilSel, 7) = Format(CCur(fgDIF.TextMatrix(FilSel, 4)) + CCur(fgDIF.TextMatrix(FilSel, 5)) + CCur(fgDIF.TextMatrix(FilSel, 6)), "#,#0.00") ' NUEVA CUOTA
fgDIF.TextMatrix(FilSel, 8) = Format(CCur(fgDIF.TextMatrix(FilSel - 1, 8)) - CCur(fgDIF.TextMatrix(FilSel, 4)), "#,#0.00")

fgDIF.Row = FilSel
fgDIF.ForeColorRow vbBlue
lnSalCap = fgDIF.TextMatrix(FilSel, 8)

fgDIF.TextMatrix(FilSel, 13) = 0
fgDIF.TextMatrix(FilSel, 12) = 0
ldFecAnt = fgDIF.TextMatrix(FilSel, 1)
lnAnioAnt = Year(ldFecAnt)
ldCuota = ldFecAnt
For I = FilSel + 1 To ContPlan - 2
    If chkFechaFija.value = 0 Then  'PERIODO FIJO
        fgDIF.TextMatrix(I, 1) = DateAdd("d", lnPlazoAprRFA, ldFecAnt)
    Else
        ldCuota = CDate(Format(Day(NuevaFecha), "00") + "/" + Format(Month(NuevaFecha), "00") + "/" + Format(lnAnioAnt + 1, "0000"))
        fgDIF.TextMatrix(I, 1) = ldCuota
    End If
    lnDiasTrans = DateDiff("D", ldFecAnt, fgDIF.TextMatrix(I, 1))
    TasaReal = InteresReal(lnTasaRFC / 100, lnDiasTrans)
    NuevoInteres = TasaReal * lnSalCap
    
    fgDIF.TextMatrix(I, 2) = lnDiasTrans
    fgDIF.TextMatrix(I, 4) = Format(CCur(fgDIF.TextMatrix(I, 7)) - NuevoInteres, "#,#0.00") 'NUEVO CAPITAL
    fgDIF.TextMatrix(I, 5) = Format(NuevoInteres, "#,#0.00") 'NUEVO INTERES
    fgDIF.TextMatrix(I, 7) = Format(CCur(fgDIF.TextMatrix(I, 4)) + CCur(fgDIF.TextMatrix(I, 5)) + CCur(fgDIF.TextMatrix(I, 6)), "#,#0.00")
    fgDIF.TextMatrix(I, 8) = Format(CCur(fgDIF.TextMatrix(I - 1, 8)) - CCur(fgDIF.TextMatrix(I, 4)), "#,#0.00")
    
    fgDIF.Row = I
    fgDIF.ForeColorRow vbBlue
    lnSalCap = fgDIF.TextMatrix(I, 8)

    fgDIF.TextMatrix(I, 13) = 0
    fgDIF.TextMatrix(I, 12) = 0
    
    ldFecAnt = fgDIF.TextMatrix(I, 1)
    lnAnioAnt = Year(ldFecAnt)
Next I
'ahora calculamos el capital para la cuota final
lnDiasTrans = DateDiff("D", ldFecAnt, fgDIF.TextMatrix(ContPlan - 1, 1))
TasaReal = InteresReal(lnTasaRFC / 100, lnDiasTrans)
NuevoInteres = TasaReal * lnSalCap
fgDIF.TextMatrix(ContPlan - 1, 2) = lnDiasTrans
fgDIF.TextMatrix(ContPlan - 1, 4) = Format(lnSalCap, "#,#0.00") 'NUEVO CAPITAL
fgDIF.TextMatrix(ContPlan - 1, 5) = Format(NuevoInteres, "#,#0.00") 'NUEVO INTERES
fgDIF.TextMatrix(ContPlan - 1, 7) = Format(CCur(fgDIF.TextMatrix(ContPlan - 1, 4)) + CCur(fgDIF.TextMatrix(ContPlan - 1, 5)) + CCur(fgDIF.TextMatrix(ContPlan - 1, 6)), "#,#0.00")
fgDIF.TextMatrix(ContPlan - 1, 8) = Format(CCur(fgDIF.TextMatrix(ContPlan - 2, 8)) - CCur(fgDIF.TextMatrix(ContPlan - 1, 4)), "#,#0.00")


fgDIF.Row = ContPlan - 1
fgDIF.ForeColorRow vbBlue
lnSalCap = fgDIF.TextMatrix(ContPlan - 1, 8)

lnSalCapNew = 0
lnSalIntNew = 0
For I = 1 To ContPlan - 1
    If fgDIF.TextMatrix(I, 9) <> "P" Then
        lnSalCapNew = lnSalCapNew + CCur(fgDIF.TextMatrix(I, 4))
        lnSalIntNew = lnSalIntNew + CCur(fgDIF.TextMatrix(I, 5))
    End If
Next

lblSalCapNewDIf = Format(lnSalCapNew, "#,#0.00")
lblSalIntNewDIF = Format(lnSalIntNew, "#,#0.00")

ReprogramaDIF = True

End Function

Function Reprogramación(ByVal FilSel As Long) As Boolean

Reprogramación = ReprogramaRFA(FilSel)
If Reprogramación = False Then Exit Function

Reprogramación = ReprogramaRFC(FilSel)
If Reprogramación = False Then Exit Function

Reprogramación = ReprogramaDIF(FilSel)
If Reprogramación = False Then Exit Function

End Function

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub cmdImpresion_Click()

    lsImpCopia = "ESTA ES UNA COPIA PREVIA A LA REESTRUCTURACION (DOCUMENTO NO VALIDO)"
    PrtCal1
    PrtCal2
    
End Sub

Private Sub Form_Load()
Set oRFA = New COMNCredito.NCOMRFA

End Sub

Function InteresReal(ByVal pnTasInt As Double, ByVal pnPeriod As Integer, Optional ByVal pnPerCuo As Integer = 360) As Double
    InteresReal = ((1 + pnTasInt) ^ (pnPeriod / pnPerCuo)) - 1
End Function

Private Sub Form_Unload(Cancel As Integer)
Set oRFA = Nothing
End Sub

Sub CargaDatos(ByVal psPersCod As String)

Dim rs As ADODB.Recordset

Set rs = oRFA.GetCreditosRFA(psPersCod)
    
If Not rs.EOF And Not rs.BOF Then
    CargaCalendariosOrig psPersCod, IIf(CHKRefinanciados.value = 1, True, False)
Else
    MsgBox "Cliente no posee creditos en RFA", vbInformation, "Aviso"
    Exit Sub
End If
rs.Close
Set rs = Nothing
End Sub

Sub CargaCalendariosOrig(ByVal psPersCod As String, ByVal pbRefinan As Boolean)

Dim lnSaldoCap As Currency
Dim lnDias As Long
Dim lnSalInt As Currency
Dim ldFecAnt As Date
Dim oRFA As COMNCredito.NCOMRFA
Set oRFA = New COMNCredito.NCOMRFA

'Set rsRFA = oRFA.GetCalendariosRFA(psPersCod, "RFA", True, pbRefinan)
Call oRFA.CargaCalendariosRFA(psPersCod, pbRefinan, rsRFA, rsRFC, rsDIF)
Set oRFA = Nothing

Me.fgRFA.Clear
Me.fgRFA.Rows = 2
Me.fgRFA.FormaCabecera
lnDias = 0
lnSaldoCap = 0
lnSalInt = 0
 
If Not rsRFA.EOF And Not rsRFA.BOF Then
    lnSaldoCap = rsRFA!nCapDes
    lnSalInt = rsRFA!nSalInt
    lnTasaRFA = rsRFA!nTasInt
    lnPlazoAprRFA = rsRFA!nDiaApr
    lnCapDesRFA = rsRFA!nCapDes
    gsCodCtaRFA = rsRFA!cCodCta
    sCtaCodRFA = rsRFA!cCodCta
    
    lblCapDesRFA = Format(rsRFA!nCapDes, "#,#0.00")
    lblnSalCapRFA = Format(rsRFA!nSaldoK, "#,#0.00")
    lblSalIntRFA = Format(lnSalInt, "#,#0.00")
    ldFecAnt = Format(rsRFA!dFecVig, "dd/mm/yyyy")
    Do While Not rsRFA.EOF
        lnDias = DateDiff("d", ldFecAnt, rsRFA!dFecVen)
        fgRFA.AdicionaFila
        fgRFA.TextMatrix(fgRFA.Row, 1) = rsRFA!dFecVen
        fgRFA.TextMatrix(fgRFA.Row, 2) = lnDias
        fgRFA.TextMatrix(fgRFA.Row, 3) = rsRFA!cNrocuo
        fgRFA.TextMatrix(fgRFA.Row, 4) = Format(rsRFA!nCapita, "#,#0.00")
        fgRFA.TextMatrix(fgRFA.Row, 5) = Format(rsRFA!nIntere, "#,#0.00")
        fgRFA.TextMatrix(fgRFA.Row, 6) = Format(rsRFA!nCofide, "#,#0.00")
        fgRFA.TextMatrix(fgRFA.Row, 7) = Format(rsRFA!nCapita + rsRFA!nIntere, "#,#0.00")
        lnSaldoCap = lnSaldoCap - rsRFA!nCapita
        fgRFA.TextMatrix(fgRFA.Row, 8) = Format(lnSaldoCap, "#,#0.00")
        fgRFA.TextMatrix(fgRFA.Row, 9) = rsRFA!cEstado
        fgRFA.TextMatrix(fgRFA.Row, 10) = rsRFA!nCapPag
        fgRFA.TextMatrix(fgRFA.Row, 11) = rsRFA!nIntPag
        fgRFA.TextMatrix(fgRFA.Row, 12) = rsRFA!nMorPag
        fgRFA.TextMatrix(fgRFA.Row, 13) = rsRFA!nIntMor
        If rsRFA!cEstado = "1" Then
            fgRFA.BackColorRow vbYellow
        End If
        ldFecAnt = Format(rsRFA!dFecVen, "dd/mm/yyyy")
        rsRFA.MoveNext
    Loop
End If


Me.fgRFC.Clear
Me.fgRFC.Rows = 2
Me.fgRFC.FormaCabecera
lnDias = 0
lnSaldoCap = 0
lnSalInt = 0
'Set rsRFC = oRFA.GetCalendariosRFA(psPersCod, "RFC", True, pbRefinan)

If Not rsRFC.EOF And Not rsRFC.BOF Then
    lnSaldoCap = rsRFC!nCapDes
    lnSalInt = rsRFC!nSalInt
    lnTasaRFC = rsRFC!nTasInt
    lnPlazoAprRFC = rsRFC!nDiaApr
    lnCapDesRFC = rsRFC!nCapDes
    gsCodCtaRFC = rsRFC!cCodCta
    sCtaCodRFC = rsRFC!cCodCta
    
    lblCapDesRFC = Format(rsRFC!nCapDes, "#,#0.00")
    lblnSalCapRFC = Format(rsRFC!nSaldoK, "#,#0.00")
    lblSalIntRFC = Format(lnSalInt, "#,#0.00")
    ldFecAnt = Format(rsRFC!dFecVig, "dd/mm/yyyy")
    Do While Not rsRFC.EOF
        lnDias = DateDiff("d", ldFecAnt, rsRFC!dFecVen)
        fgRFC.AdicionaFila
        fgRFC.TextMatrix(fgRFC.Row, 1) = rsRFC!dFecVen
        fgRFC.TextMatrix(fgRFC.Row, 2) = lnDias
        fgRFC.TextMatrix(fgRFC.Row, 3) = rsRFC!cNrocuo
        fgRFC.TextMatrix(fgRFC.Row, 4) = Format(rsRFC!nCapita, "#,#0.00")
        fgRFC.TextMatrix(fgRFC.Row, 5) = Format(rsRFC!nIntere, "#,#0.00")
        fgRFC.TextMatrix(fgRFC.Row, 6) = Format(rsRFC!nCofide, "#,#0.00")
        fgRFC.TextMatrix(fgRFC.Row, 7) = Format(rsRFC!nCapita + rsRFC!nIntere + rsRFC!nCofide, "#,#0.00")
        lnSaldoCap = lnSaldoCap - rsRFC!nCapita
        fgRFC.TextMatrix(fgRFC.Row, 8) = Format(lnSaldoCap, "#,#0.00")
        fgRFC.TextMatrix(fgRFC.Row, 9) = rsRFC!cEstado
        fgRFC.TextMatrix(fgRFC.Row, 10) = rsRFC!nCapPag
        fgRFC.TextMatrix(fgRFC.Row, 11) = rsRFC!nIntPag
        fgRFC.TextMatrix(fgRFC.Row, 12) = rsRFC!nMorPag
        fgRFC.TextMatrix(fgRFC.Row, 13) = rsRFC!nIntMor
        If rsRFC!cEstado = "P" Then
            fgRFC.BackColorRow vbYellow
        End If
        ldFecAnt = Format(rsRFC!dFecVen, "dd/mm/yyyy")
        rsRFC.MoveNext
    Loop
End If


Me.fgDIF.Clear
Me.fgDIF.Rows = 2
Me.fgDIF.FormaCabecera
lnDias = 0
lnSaldoCap = 0
lnSalInt = 0

'Set rsDIF = oRFA.GetCalendariosRFA(psPersCod, "DIF", True, pbRefinan)
If Not rsDIF.EOF And Not rsDIF.BOF Then
    lnSaldoCap = rsDIF!nCapDes
    
    lnSalInt = rsDIF!nSalInt
    lnTasaDIF = rsDIF!nTasInt
    lnPlazoAprDIF = rsDIF!nDiaApr
    lnCapDesDIF = rsDIF!nCapDes
    gsCodCtaDIF = rsDIF!cCodCta
    sCtaCodDIF = rsDIF!cCodCta
    
    lblCapDesDIF = Format(rsDIF!nCapDes, "#,#0.00")
    lblnSalCapDIF = Format(rsDIF!nSaldoK, "#,#0.00")
    lblSalIntDIF = Format(lnSalInt, "#,#0.00")
    ldFecAnt = Format(rsDIF!dFecVig, "dd/mm/yyyy")
    Do While Not rsDIF.EOF
        lnDias = DateDiff("d", ldFecAnt, rsDIF!dFecVen)
        fgDIF.AdicionaFila
        fgDIF.TextMatrix(fgDIF.Row, 1) = rsDIF!dFecVen
        fgDIF.TextMatrix(fgDIF.Row, 2) = lnDias
        fgDIF.TextMatrix(fgDIF.Row, 3) = rsDIF!cNrocuo
        fgDIF.TextMatrix(fgDIF.Row, 4) = Format(rsDIF!nCapita, "#,#0.00")
        fgDIF.TextMatrix(fgDIF.Row, 5) = Format(rsDIF!nIntere, "#,#0.00")
        fgDIF.TextMatrix(fgDIF.Row, 6) = Format(rsDIF!nCofide, "#,#0.00")
        fgDIF.TextMatrix(fgDIF.Row, 7) = Format(rsDIF!nCapita + rsDIF!nIntere + rsDIF!nCofide, "#,#0.00")
        lnSaldoCap = lnSaldoCap - rsDIF!nCapita
        fgDIF.TextMatrix(fgDIF.Row, 8) = Format(lnSaldoCap, "#,#0.00")
        fgDIF.TextMatrix(fgDIF.Row, 9) = rsDIF!cEstado
        fgDIF.TextMatrix(fgDIF.Row, 10) = rsDIF!nCapPag
        fgDIF.TextMatrix(fgDIF.Row, 11) = rsDIF!nIntPag
        fgDIF.TextMatrix(fgDIF.Row, 12) = rsDIF!nMorPag
        fgDIF.TextMatrix(fgDIF.Row, 13) = rsDIF!nIntMor
        If rsDIF!cEstado = "P" Then
            fgDIF.BackColorRow vbYellow
        End If
        ldFecAnt = Format(rsDIF!dFecVen, "dd/mm/yyyy")
        rsDIF.MoveNext
    Loop
End If

End Sub
'********************************************************************
'*FUNCION QUE HALLA EL INTERES REAL DE ACUERDO A LOS DIAS DE PLAZO
'* Ejemplo1 : TI = 4.0 a 30 dias = 4.0
'* Ejemplo2 : TI = 4.0 a 40 dias = 5.16
'********************************************************************
Function InteresReal1(ByVal Ti As Double, ByVal Periodo As Integer) As Double
    InteresReal1 = ((1 + Ti) ^ (Periodo / 360)) - 1
End Function

'***********************************************
'*  FUNCION QUE HALLA EL INTERES DE UN PERIODO DE DIAS TRANACURRIDOS
'***********************************************
Function IntPerDias1(ByVal inter As Double, ByVal DiasTrans As Integer, ByVal Periodo As Double) As Double
    IntPerDias1 = ((1 + (inter / 100)) ^ (DiasTrans / Periodo)) - 1
End Function
'*********************************************************
'FUNCION QUE HALLA LA CUOTA FIJA APARTIR DE LA TASA DE INTERES Y EL PLAZO
'*********************************************************
Function CFija(ptasa As Double, ByVal pPlazo As Integer, ByVal vMonto As Double, ByVal pper As Double) As Double
Dim pot1 As Double
    ptasa = InteresReal(ptasa, pper)
'Obtengo la cuota de pago
    pot1 = (1 + ptasa) ^ pPlazo
    CFija = ((pot1 * ptasa) / (pot1 - 1)) * vMonto
End Function

Private Sub txtBuscaCli_EmiteDatos()
If txtBuscaCli <> "" Then
    'lbOtroPag = False
    lblnomcli = txtBuscaCli.psDescripcion
    lsDocID = txtBuscaCli.sPersNroDoc
    lsDocJur = ""
    CargaDatos txtBuscaCli
Else
   CmdBuscar.SetFocus
End If
End Sub

Private Sub txtFechaNew_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdAceptar.Enabled And cmdAceptar.Visible Then cmdAceptar.SetFocus
End If
End Sub
Private Function CabPrt1() As String
    Dim lsCad As String
    Dim nLineas As Integer
    
    Linea lsCad, ImpreFormat(gsNomCmac, 60) & "Reporte al:" & FechaHora(gdFecSis)
    Linea lsCad, ImpreFormat(gsNomAge, 60) & "USUARIO : " & gsCodUser, 2
    nLineas = nLineas + 2
    Linea lsCad, PrnSet("Esp", 4) & PrnSet("B+") + CentrarCadena("CALENDARIO DE PAGO - RFA - REESTRUCTURADO", 80)
    'Linea lsCad, PrnSet("Esp", 4) & PrnSet("B+") + CentrarCadena("CALENDARIO DE PAGO - RFA ", 80)
    nLineas = nLineas + 1
    Linea lsCad, CentrarCadena(String(55, "-"), 80) + PrnSet("B-") + PrnSet("EspN"), 2
    nLineas = nLineas + 3
    If lsImpCopia <> "" Then
        Linea lsCad, PrnSet("I+") + PrnSet("B+") + CentrarCadena(lsImpCopia, 80) + PrnSet("I-"), 2
        nLineas = nLineas + 2
    End If
    
    Linea lsCad, PrnSet("C+") & ImpreFormat("CLIENTE:  " & Trim(txtBuscaCli.Text), 30, 0) & ImpreFormat("NOMBRE :  " & Trim(lblnomcli), 50)
    nLineas = nLineas + 1
    Linea lsCad, ImpreFormat("DOC.NAT.:  " & Trim(lsDocID), 30, 0) & ImpreFormat("DOC.JUR.:  " & Trim(lsDocJur), 30), 2
    nLineas = nLineas + 3
    Linea lsCad, ImpreFormat("REFINANCIADO x R.F.A.             :  " & Format(lblCapDesRFA, "#,##0.00"), 50) & ImpreFormat("TASA INTERES(ANUAL)" & Format(lnTasaRFA, "#,##0.00") & " %     ", 40) & ImpreFormat("TIPO DE CUOTA :  " & "FIJA", 50)
    nLineas = nLineas + 1
    Linea lsCad, ImpreFormat("REFINANCIADO x CMAC               :  " & Format(CCur(lblCapDesRFC) + CCur(lblCapDesDIF), "#,##0.00"), 50) & ImpreFormat("NUMERO CUOTAS PACTADO :  " & Format(fgRFA.Rows - 1, "###") & "       ", 40) & ImpreFormat("TIPO PERÍODO  :  " & IIf(chkFechaFija = 0, "FIJO", "FECHA FIJA"), 50)
    nLineas = nLineas + 1
    Linea lsCad, ImpreFormat("FECHA DE REESTRUCTURACION         :  " & Format(gdFecSis, "dd/mm/yyyy"), 50) & PrnSet("C-")
    nLineas = nLineas + 1
    Linea lsCad, PrnSet("B+") + String(71, "-")
    nLineas = nLineas + 1
    Linea lsCad, "CUOTA  " & "FECHA PAGO  " & "  CUOTA CMAC " & "   CUOTA RFA " & "COMIS.COFIDE " & "  VALOR CUOTA"
    nLineas = nLineas + 1
    Linea lsCad, String(71, "-") + PrnSet("B-")
    nLineas = nLineas + 3
    
    CabPrt1 = lsCad
End Function

Private Sub PrtCal1()
    Dim R As Byte
    Dim sCadena As String
    Dim lsReestruc As String
    Dim nLineas As Integer
    
    nLineas = 0
    sCadena = ""
    sCadena = CabPrt1
    
    If fgRFA.TextMatrix(0, 0) <> "" Then
       For R = 1 To fgRFA.Rows - 1
           If Val(fgRFA.TextMatrix(R, 3)) >= Val(lblNroCuo) Then
                lsReestruc = "(*)"
           Else
                lsReestruc = ""
           End If
           Linea sCadena, IIf(fgRFA.TextMatrix(R, 9) = "P", PrnSet("B+"), "") & ImpreFormat(Format(fgRFA.TextMatrix(R, 3), "0000"), 7, 0) + ImpreFormat(Format(fgRFA.TextMatrix(R, 1), "dd/mm/yyyy"), 12, 0) + ImpreFormat(CCur(fgRFC.TextMatrix(R, 7)) + CCur(fgDIF.TextMatrix(R, 7)), 9, 2) + " " + ImpreFormat(CCur(fgRFA.TextMatrix(R, 7)), 9, 2) + " " + _
                 ImpreFormat(CCur(fgRFA.TextMatrix(R, 6)), 9, 2) + " " + ImpreFormat(CCur(fgRFC.TextMatrix(R, 7)) + CCur(fgDIF.TextMatrix(R, 7)) + CCur(fgRFA.TextMatrix(R, 7)) + CCur(fgRFA.TextMatrix(R, 6)), 9, 2) & " " & IIf(fgRFA.TextMatrix(R, 9) = "P", "PAG", lsReestruc) & PrnSet("B-")
                nLineas = nLineas + 1
           
           If nLineas > 60 Then
              nLineas = 0
              sCadena = sCadena & Chr$(12)
              sCadena = sCadena + CabPrt1
           End If
       Next R
       Linea sCadena, PrnSet("B+") + String(71, "-")
       nLineas = nLineas + 1
       Linea sCadena, "Copia de Calendario de Pagos para cliente acogido al programa R.F.A." + PrnSet("B-") + PrnSet("C-")
       Linea sCadena, PrnSet("Ss") + PrnSet("12CPI")
       EnviaPrevio sCadena, "Calendario Consolidado Reestructurado", 66
    End If
End Sub

Private Function CabPrt2() As String
Dim lsCad As String
Dim nLineas As Integer

    Linea lsCad, ImpreFormat(gsNomCmac, 60) & "Reporte al:" & FechaHora(gdFecSis)
    Linea lsCad, ImpreFormat(gsNomAge, 60) & "USUARIO : " & gsCodUser, 2
    nLineas = nLineas + 2
    Linea lsCad, PrnSet("Esp", 4) & PrnSet("B+") + CentrarCadena("CALENDARIO DE PAGO - RFA - REESTRUCTURADO", 80)
    'Linea lsCad, PrnSet("Esp", 4) & PrnSet("B+") + CentrarCadena("CALENDARIO DE PAGO - RFA", 80)
    nLineas = nLineas + 1
    Linea lsCad, CentrarCadena(String(55, "-"), 80) + PrnSet("B-") + PrnSet("EspN"), 2
    nLineas = nLineas + 3
    If lsImpCopia <> "" Then
        Linea lsCad, PrnSet("I+") + PrnSet("B+") + CentrarCadena(lsImpCopia, 80) + PrnSet("I-"), 2
        nLineas = nLineas + 2
    End If
    Linea lsCad, PrnSet("C+") & ImpreFormat("CLIENTE:  " & Trim(txtBuscaCli.Text), 30, 0) & ImpreFormat("NOMBRE :  " & Trim(lblnomcli), 50)
    nLineas = nLineas + 1
    Linea lsCad, ImpreFormat("DOC.NAT.:  " & Trim(lsDocID), 30, 0) & ImpreFormat("DOC.JUR.:  " & Trim(lsDocJur), 30), 2
    nLineas = nLineas + 3
    Linea lsCad, ImpreFormat("REFINANCIADO x R.F.A.             :  " & Format(lblCapDesRFA, "#,##0.00"), 50) & ImpreFormat("TASA INTERES(ANUAL)" & Format(lnTasaRFA, "#,##0.00") & " %     ", 40) & ImpreFormat("TIPO DE CUOTA :  " & "FIJA", 50)
    nLineas = nLineas + 1
    Linea lsCad, ImpreFormat("REFINANCIADO x CMAC               :  " & Format(CCur(lblCapDesRFC) + CCur(lblCapDesDIF), "#,##0.00"), 50) & ImpreFormat("NUMERO CUOTAS PACTADO :  " & Format(fgRFA.Rows - 1, "###") & "       ", 40) & ImpreFormat("TIPO PERÍODO  :  " & IIf(chkFechaFija = 0, "FIJO", "FECHA FIJA"), 50)
    nLineas = nLineas + 1
    Linea lsCad, ImpreFormat("FECHA DE REESTRUCTURACION         :  " & Format(gdFecSis, "dd/mm/yyyy"), 50) & PrnSet("C-")
    nLineas = nLineas + 1
    Linea lsCad, PrnSet("C+") + PrnSet("B+") + String(120, "-")
    nLineas = nLineas + 1
    Linea lsCad, "CUOT " & "FECHA PAGO " & "DIAS " & "CAP. CMAC " & "INT. CMAC " & "CUOTA.CMAC " & "CAP.DIF." & "INT.DIF." & "CUOTA.DIF" & "CAP.  RFA " & "INT.  RFA " & "IN.G. RFA " & "CUOTA RFA " & "C. COFIDE " & "VAL.CUOTA "
    nLineas = nLineas + 1
    Linea lsCad, String(120, "-") + PrnSet("C-") + PrnSet("B-")
    nLineas = nLineas + 3
    
    CabPrt2 = lsCad
End Function

Private Sub PrtCal2()
    Dim R As Byte
    Dim lsReestruc As String
    Dim sCadena As String
    Dim nLineas As Integer
    
    nLineas = 0
    sCadena = ""
    sCadena = CabPrt2
  
    If fgRFA.TextMatrix(0, 0) <> "" Then
       For R = 1 To fgRFA.Rows - 1
            If Val(fgRFA.TextMatrix(R, 3)) >= Val(lblNroCuo) Then
                lsReestruc = "(*)"
            Else
                lsReestruc = ""
            End If
           
           Linea sCadena, PrnSet("C+") + IIf(fgRFA.TextMatrix(R, 9) = "P", PrnSet("B+"), "") + ImpreFormat(Format(fgRFA.TextMatrix(R, 3), "0000"), 5, 0) + ImpreFormat(Format(fgRFA.TextMatrix(R, 1), "dd/mm/yyyy"), 11, 0) + _
                 ImpreFormat(Format(fgRFA.TextMatrix(R, 2), "0000"), 5, 0) + ImpreFormat(CCur(fgRFC.TextMatrix(R, 4)), 6, 2) + " " + ImpreFormat(CCur(fgRFC.TextMatrix(R, 5)), 6, 2) + " " + ImpreFormat(CCur(fgRFC.TextMatrix(R, 7)), 6, 2) + " " + _
                 ImpreFormat(CCur(fgDIF.TextMatrix(R, 7)), 6, 2) + " " + ImpreFormat(CCur(fgRFA.TextMatrix(R, 4)), 6, 2) + " " + ImpreFormat(CCur(fgRFA.TextMatrix(R, 5)), 6, 2) + " " + ImpreFormat(0, 6, 2) + " " + ImpreFormat(CCur(fgRFA.TextMatrix(R, 4)) + CCur(fgRFA.TextMatrix(R, 5)), 6, 2) + " " + _
                 ImpreFormat(CCur(fgRFA.TextMatrix(R, 6)), 6, 2) + " " + ImpreFormat(CCur(fgRFC.TextMatrix(R, 7)) + CCur(fgDIF.TextMatrix(R, 7)) + CCur(fgRFA.TextMatrix(R, 7)) + CCur(fgRFA.TextMatrix(R, 6)), 6, 2) + PrnSet("C-") & " " & IIf(fgRFA.TextMatrix(R, 9) = "P", "PAG", lsReestruc) & PrnSet("B-")
           nLineas = nLineas + 1
           If nLineas > 60 Then
              nLineas = 0
              sCadena = sCadena & Chr$(12)
              sCadena = sCadena + CabPrt1
           End If
       Next R
       Linea sCadena, PrnSet("C+") + PrnSet("B+") + String(120, "-")
'       nLineas = nLineas + 1
       Linea sCadena, "NUEVOS SALDOS REESTRUCTURADOS "
       Linea sCadena, "Capital RFC  " & ImpreFormat(CCur(lblSalCapNewRFC) + CCur(lblSalCapNewDIf), 7, 2) & Space(20) & "Capital RFA  " & ImpreFormat(CCur(lblSalCapNewRFA), 7, 2)
       Linea sCadena, "Interes RFC  " & ImpreFormat(CCur(lblSalIntNewRFC) + CCur(lblSalIntNewDIF), 7, 2) & Space(20) & "Interes RFA  " & ImpreFormat(CCur(lblSalIntNewRFA), 7, 2)
       
       Linea sCadena, "Total   RFC  " & ImpreFormat(CCur(lblSalCapNewRFC) + CCur(lblSalCapNewDIf) + CCur(lblSalIntNewRFC) + CCur(lblSalIntNewDIF), 7, 2) & Space(20) & "Total   RFA  " & ImpreFormat(CCur(lblSalCapNewRFA) + CCur(lblSalIntNewRFA), 7, 2)
       
       Linea sCadena, PrnSet("C+") + PrnSet("B+") + String(120, "-")
       
       Linea sCadena, "Este calendario es de uso exclusivo para el responsable del Programa R.F.A. - utilizar para Contrato"
       Linea sCadena, PrnSet("Ss") + PrnSet("12CPI")
    
        EnviaPrevio sCadena, "Calendario Consolidado Reestructurado", 66
    End If
    
End Sub


Sub DistribuyeMatrices()
Dim I As Integer
    'rfa
    ReDim MatCalendRFC(fgRFA.Rows - 1, 11)
    ReDim MatCalendDIF(fgRFC.Rows - 1, 11)
    ReDim MatCalendRFA(fgDIF.Rows - 1, 11)
    
    For I = 1 To fgRFA.Rows - 1
         MatCalendRFA(I, 0) = sCtaCodRFA
         MatCalendRFA(I, 1) = fgRFA.TextMatrix(I, 3) 'cuota
         MatCalendRFA(I, 2) = fgRFA.TextMatrix(I, 1) 'Fecha Vencimiento
         MatCalendRFA(I, 3) = fgRFA.TextMatrix(I, 9) 'Estado
         MatCalendRFA(I, 4) = fgRFA.TextMatrix(I, 4) 'capital
         MatCalendRFA(I, 5) = fgRFA.TextMatrix(I, 5) 'Interes
         MatCalendRFA(I, 6) = fgRFA.TextMatrix(I, 6) 'Cofide
         MatCalendRFA(I, 7) = fgRFA.TextMatrix(I, 13) 'Mora
         MatCalendRFA(I, 8) = fgRFA.TextMatrix(I, 10) 'Capital Pagado
         MatCalendRFA(I, 9) = fgRFA.TextMatrix(I, 11) 'Interes Pagado
         MatCalendRFA(I, 10) = fgRFA.TextMatrix(I, 12) 'Interes Pagado
    Next I
    
    For I = 1 To fgDIF.Rows - 1
         MatCalendDIF(I, 0) = sCtaCodDIF
         MatCalendDIF(I, 1) = fgDIF.TextMatrix(I, 3) 'cuota
         MatCalendDIF(I, 2) = fgDIF.TextMatrix(I, 1) 'Fecha Vencimiento
         MatCalendDIF(I, 3) = fgDIF.TextMatrix(I, 9) 'Estado
         MatCalendDIF(I, 4) = fgDIF.TextMatrix(I, 4) 'capital
         MatCalendDIF(I, 5) = fgDIF.TextMatrix(I, 5) 'Interes
         MatCalendDIF(I, 6) = fgDIF.TextMatrix(I, 6) 'Cofide
         MatCalendDIF(I, 7) = fgDIF.TextMatrix(I, 13) 'Mora
         MatCalendDIF(I, 8) = fgDIF.TextMatrix(I, 10) 'Capital Pagado
         MatCalendDIF(I, 9) = fgDIF.TextMatrix(I, 11) 'Interes Pagado
         MatCalendDIF(I, 10) = fgDIF.TextMatrix(I, 12) 'Interes Pagado
    Next I
    
    For I = 1 To fgRFC.Rows - 1
         MatCalendRFC(I, 0) = sCtaCodRFC
         MatCalendRFC(I, 1) = fgRFC.TextMatrix(I, 3) 'cuota
         MatCalendRFC(I, 2) = fgRFC.TextMatrix(I, 1) 'Fecha Vencimiento
         MatCalendRFC(I, 3) = fgRFC.TextMatrix(I, 9) 'Estado
         MatCalendRFC(I, 4) = fgRFC.TextMatrix(I, 4) 'capital
         MatCalendRFC(I, 5) = fgRFC.TextMatrix(I, 5) 'Interes
         MatCalendRFC(I, 6) = fgRFC.TextMatrix(I, 6) 'Cofide
         MatCalendRFC(I, 7) = fgRFC.TextMatrix(I, 13) 'Mora
         MatCalendRFC(I, 8) = fgRFC.TextMatrix(I, 10) 'Capital Pagado
         MatCalendRFC(I, 9) = fgRFC.TextMatrix(I, 11) 'Interes Pagado
         MatCalendRFC(I, 10) = fgRFC.TextMatrix(I, 12) 'Interes Pagado
    Next I
    
End Sub
