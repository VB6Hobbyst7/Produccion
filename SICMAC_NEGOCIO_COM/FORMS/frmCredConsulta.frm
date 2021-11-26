VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Credito"
   ClientHeight    =   6945
   ClientLeft      =   1455
   ClientTop       =   2655
   ClientWidth     =   11760
   Icon            =   "frmCredConsulta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDeudaCuoIFIS_CD 
      Caption         =   "&Deuda IFIs"
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
      Left            =   5400
      TabIndex        =   172
      Top             =   6405
      Width           =   1845
   End
   Begin VB.Frame Frame10 
      Height          =   2040
      Left            =   480
      TabIndex        =   113
      Top             =   0
      Width           =   10635
      Begin SICMACT.ActXCodCta ActxCta 
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
         TabIndex        =   114
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
         Caption         =   "Tipo de Producto :"
         Height          =   195
         Index           =   36
         Left            =   120
         TabIndex        =   147
         Top             =   1080
         Width           =   1320
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTipoProducto 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1485
         TabIndex        =   146
         Top             =   1080
         Width           =   3720
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
         TabIndex        =   118
         Top             =   1440
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         Left            =   1485
         TabIndex        =   117
         Top             =   1455
         Width           =   3705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Crédito :"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   116
         Top             =   720
         Width           =   1260
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbltipoCredito 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1485
         TabIndex        =   115
         Top             =   720
         Width           =   3720
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
      Left            =   10380
      TabIndex        =   3
      Top             =   6405
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
      Left            =   7290
      TabIndex        =   1
      Top             =   6405
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
      Left            =   9165
      TabIndex        =   2
      Top             =   6405
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4170
      Left            =   45
      TabIndex        =   4
      Top             =   2190
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7355
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      Tab             =   4
      TabsPerRow      =   9
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
      TabPicture(0)   =   "frmCredConsulta.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdIdentificacionRCC"
      Tab(0).Control(1)=   "CmdHistorial"
      Tab(0).Control(2)=   "CmdMuestra"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Historial "
      TabPicture(1)   =   "frmCredConsulta.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_Reprogramado"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Desembolsos"
      TabPicture(2)   =   "frmCredConsulta.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2(1)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Pagos &Realizados"
      TabPicture(3)   =   "frmCredConsulta.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2(2)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Pagos &Pendientes"
      TabPicture(4)   =   "frmCredConsulta.frx":007C
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label8"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "lblMonedaP"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "lblMensajeMIVIVIENDA"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "lblLiquidacion"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Frame6"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Frame7"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Frame8"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "&Garantías"
      TabPicture(5)   =   "frmCredConsulta.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fragarantias"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&Otros Datos"
      TabPicture(6)   =   "frmCredConsulta.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame2(3)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Com Pago"
      TabPicture(7)   =   "frmCredConsulta.frx":00D0
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame13"
      Tab(7).Control(1)=   "Frame12"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Visitas"
      TabPicture(8)   =   "frmCredConsulta.frx":00EC
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame14"
      Tab(8).ControlCount=   1
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
         Height          =   2685
         Left            =   9360
         TabIndex        =   200
         Top             =   450
         Width           =   1995
         Begin VB.Label lblICV 
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
            Left            =   930
            TabIndex        =   212
            Top             =   1770
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
            Left            =   930
            TabIndex        =   211
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
            Left            =   930
            TabIndex        =   210
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
            Left            =   930
            TabIndex        =   209
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Saldo K"
            Height          =   195
            Index           =   48
            Left            =   120
            TabIndex        =   208
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Int Comp"
            Height          =   195
            Index           =   46
            Left            =   120
            TabIndex        =   207
            Top             =   720
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gasto"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   206
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
            Left            =   930
            TabIndex        =   205
            Top             =   2235
            Width           =   930
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
            Left            =   930
            TabIndex        =   204
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
            TabIndex        =   203
            Top             =   2295
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Int Morat"
            Height          =   195
            Index           =   34
            Left            =   120
            TabIndex        =   202
            Top             =   1470
            Width           =   630
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Caption         =   "Int. Comp. Vencido"
            Height          =   405
            Left            =   90
            TabIndex        =   201
            Top             =   1740
            Width           =   735
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Historial de Visitas realizadas por los Gestores"
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
         Height          =   3135
         Left            =   -74880
         TabIndex        =   162
         Top             =   480
         Width           =   11415
         Begin SICMACT.FlexEdit FeAdj 
            Height          =   2775
            Left            =   120
            TabIndex        =   163
            Top             =   240
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   4895
            Cols0           =   10
            HighLight       =   1
            AllowUserResizing=   1
            RowSizingMode   =   1
            EncabezadosNombres=   "Nº-Fecha Visita-Lugar Contacto-Pers. Contac.-Resul. Gestion-Fec. Comprom.-Monto Compr.-Comentario-Registro-nMovNro"
            EncabezadosAnchos=   "400-1200-1200-1200-1200-1200-1200-2200-2000-0"
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-L-L-R-L-L-C"
            FormatosEdit    =   "0-0-0-0-0-5-2-0-0-0"
            AvanceCeldas    =   1
            TextArray0      =   "Nº"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame13 
         Height          =   735
         Left            =   -74040
         TabIndex        =   156
         Top             =   2520
         Width           =   9255
         Begin VB.CommandButton cmdCompGuardar 
            Caption         =   "Guardar"
            Height          =   375
            Left            =   7920
            TabIndex        =   157
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Fecha de Compromiso de Pago"
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
         Height          =   2055
         Left            =   -74040
         TabIndex        =   148
         Top             =   480
         Width           =   9255
         Begin MSMask.MaskEdBox txtCompFecha 
            Height          =   315
            Left            =   7440
            TabIndex        =   160
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label12 
            Caption         =   "Fecha de Compromiso de Pago:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4440
            TabIndex        =   161
            Top             =   420
            Width           =   2775
         End
         Begin VB.Label lblCompDias 
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7440
            TabIndex        =   159
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblCompDiasMsg 
            Caption         =   "Nro Dias Faltantes:"
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
            Left            =   5400
            TabIndex        =   158
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label lblCompMsg 
            Caption         =   "."
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
            Height          =   255
            Left            =   240
            TabIndex        =   155
            Top             =   1560
            Width           =   8655
         End
         Begin VB.Label lblCompCuotaFecha 
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            TabIndex        =   154
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "Fecha de Pago:"
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
            Left            =   360
            TabIndex        =   153
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblCompCuotaSaldo 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            TabIndex        =   152
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Saldo de Cuota:"
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
            Left            =   360
            TabIndex        =   151
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblCompCuota 
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            TabIndex        =   150
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "Nro. de Cuota:"
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
            Left            =   360
            TabIndex        =   149
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdIdentificacionRCC 
         Caption         =   "Identificacion RCC"
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
         Left            =   -72420
         TabIndex        =   136
         Top             =   3480
         Width           =   2115
      End
      Begin VB.CommandButton CmdHistorial 
         Caption         =   "Kardex Historial C."
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
         Left            =   -74460
         TabIndex        =   131
         Top             =   3480
         Width           =   2025
      End
      Begin VB.CommandButton CmdMuestra 
         Caption         =   "Kardex"
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
         Left            =   -65220
         TabIndex        =   130
         Top             =   3510
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Height          =   2670
         Left            =   -74520
         TabIndex        =   85
         Top             =   585
         Width           =   10365
         Begin VB.TextBox txtFechaCancelacion 
            Height          =   315
            Left            =   1500
            TabIndex        =   135
            Top             =   2070
            Width           =   1395
         End
         Begin VB.Label lblMonedaH 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "SOLES"
            Height          =   330
            Left            =   8460
            TabIndex        =   138
            Top             =   2160
            Width           =   1635
         End
         Begin VB.Label Label7 
            Caption         =   "Moneda:"
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
            Left            =   7605
            TabIndex        =   137
            Top             =   2160
            Width           =   780
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Cancelacion:"
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
            Left            =   270
            TabIndex        =   134
            Top             =   2010
            Width           =   1095
         End
         Begin VB.Label lblMontoSolicitado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2910
            TabIndex        =   119
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
            TabIndex        =   112
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
            TabIndex        =   111
            Top             =   1575
            Width           =   945
         End
         Begin VB.Label lblMontoAprobado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2910
            TabIndex        =   110
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
            TabIndex        =   109
            Top             =   1095
            Width           =   945
         End
         Begin VB.Label lblMontosugerido 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   2910
            TabIndex        =   108
            Top             =   1065
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Monto "
            Height          =   195
            Index           =   20
            Left            =   3240
            TabIndex        =   107
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblcuotasSolicitud 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4170
            TabIndex        =   106
            Top             =   570
            Width           =   675
         End
         Begin VB.Label lblcuotasAprobado 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4170
            TabIndex        =   105
            Top             =   1545
            Width           =   675
         End
         Begin VB.Label lblcuotasugerida 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4170
            TabIndex        =   104
            Top             =   1065
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº Cuotas"
            Height          =   195
            Index           =   21
            Left            =   4200
            TabIndex        =   103
            Top             =   255
            Width           =   720
         End
         Begin VB.Label lblPlazoSolicitud 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4890
            TabIndex        =   102
            Top             =   570
            Width           =   735
         End
         Begin VB.Label lblPlazoAprobado 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4890
            TabIndex        =   101
            Top             =   1545
            Width           =   735
         End
         Begin VB.Label lblPlazoSugerido 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4890
            TabIndex        =   100
            Top             =   1065
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Plazo"
            Height          =   195
            Index           =   22
            Left            =   5055
            TabIndex        =   99
            Top             =   255
            Width           =   390
         End
         Begin VB.Label lblmontoCuotaAprobada 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5655
            TabIndex        =   98
            Top             =   1545
            Width           =   1035
         End
         Begin VB.Label lblmontoCuotaSugerida 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   5655
            TabIndex        =   97
            Top             =   1050
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuotas"
            Height          =   195
            Index           =   23
            Left            =   5895
            TabIndex        =   96
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblfechsolicitud 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1515
            TabIndex        =   95
            Top             =   570
            Width           =   1365
         End
         Begin VB.Label lblfechaAprobado 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1515
            TabIndex        =   94
            Top             =   1545
            Width           =   1365
         End
         Begin VB.Label lblfechasugerida 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1515
            TabIndex        =   93
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
            TabIndex        =   92
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblGraciaAprobada 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6705
            TabIndex        =   91
            Top             =   1545
            Width           =   615
         End
         Begin VB.Label lblGraciasugerida 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   6705
            TabIndex        =   90
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
            TabIndex        =   89
            Top             =   240
            Width           =   765
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblTipoGraciaApr 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   555
            Left            =   8415
            TabIndex        =   88
            Top             =   1455
            Width           =   1740
         End
         Begin VB.Label lblIntGraciaApr 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   7425
            TabIndex        =   87
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
            TabIndex        =   86
            Top             =   255
            Width           =   1155
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Cuotas Pendiente:"
         Height          =   3045
         Left            =   240
         TabIndex        =   73
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
            TabIndex        =   74
            Top             =   2175
            Width           =   6660
            Begin VB.Label lblIntCompVenc 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   4350
               TabIndex        =   185
               Top             =   375
               Width           =   945
            End
            Begin VB.Label IntCompVencido 
               AutoSize        =   -1  'True
               Caption         =   "Int. Vencido"
               Height          =   195
               Index           =   58
               Left            =   4350
               TabIndex        =   184
               Top             =   180
               Width           =   855
            End
            Begin VB.Label lblCapitalCuoPend 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   150
               TabIndex        =   84
               Top             =   375
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Capital"
               Height          =   195
               Index           =   54
               Left            =   150
               TabIndex        =   83
               Top             =   180
               Width           =   480
            End
            Begin VB.Label lblGastoCuoPend 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   3360
               TabIndex        =   82
               Top             =   375
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Mora/Penalidad"
               Height          =   195
               Index           =   55
               Left            =   2055
               TabIndex        =   81
               Top             =   165
               Width           =   1140
            End
            Begin VB.Label lblMoraCuoPend 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   2040
               TabIndex        =   80
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Gasto"
               Height          =   195
               Index           =   56
               Left            =   3360
               TabIndex        =   79
               Top             =   165
               Width           =   420
            End
            Begin VB.Label lblTotalCuoPend 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   5490
               TabIndex        =   78
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
               Left            =   5490
               TabIndex        =   77
               Top             =   165
               Width           =   510
            End
            Begin VB.Label lblInteresCuoPend 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   1095
               TabIndex        =   76
               Top             =   375
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Interes "
               Height          =   195
               Index           =   60
               Left            =   1110
               TabIndex        =   75
               Top             =   165
               Width           =   525
            End
         End
         Begin MSComctlLib.ListView lstCuotasPend 
            Height          =   1965
            Left            =   57
            TabIndex        =   183
            Top             =   210
            Width           =   6735
            _ExtentX        =   11880
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
            NumItems        =   15
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Cuota"
               Object.Width           =   1587
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
               Text            =   "Int.Com.Venc"
               Object.Width           =   2647
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Mora/Penalidad"
               Object.Width           =   2647
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Gastos"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Gast. y Comi."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Seguro Desgr."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Text            =   "Seg. Contra Incendio"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Text            =   "Seg. Vehicular"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Text            =   "ITF"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   12
               Text            =   "Monto Cuota"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   13
               Text            =   "Atraso"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   14
               Text            =   "Saldo Cap."
               Object.Width           =   1941
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3135
         Index           =   2
         Left            =   -74520
         TabIndex        =   72
         Top             =   450
         Width           =   10695
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
            Height          =   2565
            Left            =   8010
            TabIndex        =   186
            Top             =   180
            Width           =   2550
            Begin VB.Label lblcapitalpagado 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1395
               TabIndex        =   198
               Top             =   450
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Capital :"
               Height          =   195
               Index           =   30
               Left            =   120
               TabIndex        =   197
               Top             =   480
               Width           =   570
            End
            Begin VB.Label lblintcompPag 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1395
               TabIndex        =   196
               Top             =   780
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Interés :"
               Height          =   195
               Index           =   31
               Left            =   120
               TabIndex        =   195
               Top             =   825
               Width           =   570
            End
            Begin VB.Label lblIntMorPag 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1395
               TabIndex        =   194
               Top             =   1425
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Mora/Penalidad :"
               Height          =   195
               Index           =   32
               Left            =   120
               TabIndex        =   193
               Top             =   1470
               Width           =   1230
            End
            Begin VB.Label lblGastopagado 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1395
               TabIndex        =   192
               Top             =   1110
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Gasto :"
               Height          =   195
               Index           =   33
               Left            =   120
               TabIndex        =   191
               Top             =   1140
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "ITF Acum:"
               Height          =   195
               Index           =   8
               Left            =   120
               TabIndex        =   190
               Top             =   2160
               Width           =   735
            End
            Begin VB.Label lblITFPagado 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1395
               TabIndex        =   189
               Top             =   2100
               Width           =   960
            End
            Begin VB.Label lblIntCompVencido 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1395
               TabIndex        =   188
               Top             =   1770
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Int.Comp.Venc."
               Height          =   195
               Index           =   49
               Left            =   120
               TabIndex        =   187
               Top             =   1830
               Width           =   1095
            End
         End
         Begin MSComctlLib.ListView ListaPagos 
            Height          =   2760
            Left            =   120
            TabIndex        =   199
            Top             =   210
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   4868
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
            NumItems        =   16
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
               Text            =   "Mora/Penalidad"
               Object.Width           =   2647
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Int.Comp.Vencido"
               Object.Width           =   2647
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Gastos "
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Gast. y Comi."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Text            =   "Seguro Desgr."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Text            =   "Seg. Contra Incendio"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Text            =   "Seg. Vehicular"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "ITF"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   13
               Text            =   "Atraso"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "Saldo Cap"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Text            =   "Usuario"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2610
         Index           =   1
         Left            =   -74400
         TabIndex        =   64
         Top             =   450
         Width           =   10230
         Begin MSComctlLib.ListView listaDesembolsos 
            Height          =   2175
            Left            =   120
            TabIndex        =   65
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
         Begin VB.Label lbl_MontoFinanciado_lbl 
            Caption         =   "Monto Financiado : "
            Height          =   255
            Left            =   7200
            TabIndex        =   167
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lbl_montoFinanciado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8760
            TabIndex        =   166
            Top             =   2040
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Desembolso :"
            Height          =   210
            Index           =   50
            Left            =   7170
            TabIndex        =   71
            Top             =   1350
            Width           =   1365
            WordWrap        =   -1  'True
         End
         Begin VB.Label lbltotalDesembolso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8775
            TabIndex        =   70
            Top             =   1305
            Width           =   1155
         End
         Begin VB.Label lblmontoDesembolsado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8775
            TabIndex        =   69
            Top             =   930
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desembolsado :"
            Height          =   195
            Index           =   28
            Left            =   7170
            TabIndex        =   68
            Top             =   990
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label lbltipoDesembolso 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   480
            Left            =   8550
            TabIndex        =   67
            Top             =   345
            Width           =   1545
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Desembolso :"
            Height          =   210
            Index           =   26
            Left            =   7170
            TabIndex        =   66
            Top             =   360
            Width           =   1320
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2850
         Left            =   -74460
         TabIndex        =   42
         Top             =   600
         Width           =   10380
         Begin VB.Label lblHonrado 
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
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1680
            TabIndex        =   182
            Top             =   2460
            Width           =   3945
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Honrado:"
            Height          =   195
            Left            =   120
            TabIndex        =   181
            Top             =   2520
            Width           =   660
         End
         Begin VB.Label LblAntigua 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   1680
            TabIndex        =   133
            Top             =   2190
            Width           =   2925
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cred.Ant."
            Height          =   195
            Left            =   120
            TabIndex        =   132
            Top             =   2190
            Width           =   660
         End
         Begin VB.Label Label5 
            Caption         =   "Com Venc :"
            Height          =   240
            Left            =   8505
            TabIndex        =   129
            Top             =   1755
            Width           =   885
         End
         Begin VB.Label LblIntComVen 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.0000"
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
            Height          =   255
            Left            =   9435
            TabIndex        =   128
            Top             =   1740
            Width           =   780
         End
         Begin VB.Label LblIntMor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.0000"
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
            Height          =   255
            Left            =   7605
            TabIndex        =   127
            Top             =   2055
            Width           =   780
         End
         Begin VB.Label Label4 
            Caption         =   "Tasa Moratoria :"
            Height          =   240
            Left            =   6090
            TabIndex        =   126
            Top             =   2085
            Width           =   1290
         End
         Begin VB.Label LblIntCom 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.0000"
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
            Height          =   255
            Left            =   7605
            TabIndex        =   125
            Top             =   1740
            Width           =   780
         End
         Begin VB.Label Label2 
            Caption         =   "Tasa Comp :"
            Height          =   240
            Left            =   6090
            TabIndex        =   124
            Top             =   1770
            Width           =   1035
         End
         Begin VB.Label lblfechavigencia 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7590
            TabIndex        =   63
            Top             =   1365
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Vigencia :"
            Height          =   195
            Index           =   15
            Left            =   6090
            TabIndex        =   62
            Top             =   1395
            Width           =   1425
         End
         Begin VB.Label lbltipocuota 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   450
            Left            =   7590
            TabIndex        =   61
            Top             =   825
            Width           =   1965
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cuota :"
            Height          =   195
            Index           =   13
            Left            =   6075
            TabIndex        =   60
            Top             =   855
            Width           =   1095
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
            Caption         =   "Destino del Crédito :"
            Height          =   195
            Index           =   12
            Left            =   105
            TabIndex        =   58
            Top             =   1905
            Width           =   1425
         End
         Begin VB.Label lblcondicion 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   57
            Top             =   1545
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Condición de Crédito :"
            Height          =   195
            Index           =   11
            Left            =   105
            TabIndex        =   56
            Top             =   1560
            Width           =   1560
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   10
            Left            =   8325
            TabIndex        =   55
            Top             =   555
            Width           =   120
         End
         Begin VB.Label lbltasainteres 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7590
            TabIndex        =   54
            Top             =   510
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tasa de Interes :"
            Height          =   195
            Index           =   9
            Left            =   6075
            TabIndex        =   53
            Top             =   555
            Width           =   1200
         End
         Begin VB.Label lblapoderado 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   52
            Top             =   1215
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Apoderado :"
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   51
            Top             =   1230
            Width           =   1005
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblnota1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7590
            TabIndex        =   50
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nota :"
            Height          =   195
            Index           =   7
            Left            =   6120
            TabIndex        =   49
            Top             =   225
            Width           =   435
         End
         Begin VB.Label lblanalista 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   48
            Top             =   885
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Analista :"
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   47
            Top             =   915
            Width           =   645
         End
         Begin VB.Label lblLinea 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   46
            Top             =   555
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Linea de Crédito :"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   45
            Top             =   600
            Width           =   1500
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblfuente 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   44
            Top             =   225
            Width           =   3945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fuente de Ingreso :"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   43
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
         Height          =   2685
         Left            =   7200
         TabIndex        =   31
         Top             =   450
         Width           =   1935
         Begin VB.Label Label14 
            Caption         =   "Penalidad"
            Height          =   255
            Left            =   90
            TabIndex        =   164
            Top             =   1620
            Width           =   735
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
            Left            =   930
            TabIndex        =   41
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Saldo K"
            Height          =   195
            Index           =   16
            Left            =   90
            TabIndex        =   40
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
            Left            =   930
            TabIndex        =   39
            Top             =   720
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Int Comp"
            Height          =   195
            Index           =   51
            Left            =   90
            TabIndex        =   38
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
            Left            =   930
            TabIndex        =   37
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gasto "
            Height          =   195
            Index           =   53
            Left            =   90
            TabIndex        =   36
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
            Left            =   930
            TabIndex        =   35
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Int Morat /"
            Height          =   195
            Index           =   57
            Left            =   90
            TabIndex        =   34
            Top             =   1420
            Width           =   750
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
            TabIndex        =   33
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
            TabIndex        =   32
            Top             =   2250
            Width           =   915
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3705
         Index           =   3
         Left            =   -74880
         TabIndex        =   7
         Top             =   360
         Width           =   11385
         Begin VB.Frame fr_SegRiesgo 
            Caption         =   "Segmentación Nivel de Riesgo "
            Height          =   735
            Left            =   120
            TabIndex        =   174
            Top             =   2903
            Width           =   11175
            Begin VB.TextBox txtMotivo 
               Enabled         =   0   'False
               Height          =   525
               Left            =   4800
               MultiLine       =   -1  'True
               TabIndex        =   180
               Top             =   150
               Width           =   6255
            End
            Begin VB.TextBox txtNivel 
               Enabled         =   0   'False
               Height          =   285
               Left            =   2640
               TabIndex        =   179
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox txtFechaCierre 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1080
               TabIndex        =   178
               Top             =   240
               Width           =   975
            End
            Begin VB.Label lblSegRisgMotivo 
               Caption         =   "Motivo"
               Height          =   255
               Left            =   4200
               TabIndex        =   177
               Top             =   240
               Width           =   615
            End
            Begin VB.Label lblSegRisgNivel 
               Caption         =   "Nivel"
               Height          =   255
               Left            =   2160
               TabIndex        =   176
               Top             =   240
               Width           =   495
            End
            Begin VB.Label lblSegRisgFC 
               Caption         =   "Fecha Cierre"
               Height          =   255
               Left            =   120
               TabIndex        =   175
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.TextBox txtGlosa 
            Height          =   300
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   169
            Top             =   720
            Width           =   4095
         End
         Begin VB.CommandButton cmdMotivoRefinanciado 
            Caption         =   "Ver Motivo Refinanciado"
            Height          =   300
            Left            =   8880
            TabIndex        =   168
            Top             =   1560
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Frame Frame11 
            Enabled         =   0   'False
            Height          =   570
            Left            =   120
            TabIndex        =   120
            Top             =   2280
            Width           =   4920
            Begin VB.CheckBox ChkCalDin 
               Caption         =   "Calend. Dinamico"
               Height          =   255
               Left            =   3255
               TabIndex        =   123
               Top             =   210
               Width           =   1560
            End
            Begin VB.CheckBox ChkMiViv 
               Caption         =   "Mi ViVienda"
               Height          =   255
               Left            =   1965
               TabIndex        =   122
               Top             =   210
               Width           =   1635
            End
            Begin VB.CheckBox ChkCuotaCom 
               Caption         =   "Cuota Comodin"
               Height          =   255
               Left            =   180
               TabIndex        =   121
               Top             =   210
               Width           =   1635
            End
         End
         Begin VB.Frame Frame3 
            Enabled         =   0   'False
            Height          =   1290
            Left            =   6240
            TabIndex        =   16
            Top             =   240
            Width           =   1500
            Begin VB.CheckBox chkRefinanciado 
               Alignment       =   1  'Right Justify
               Caption         =   "Refinanciado?"
               Height          =   285
               Left            =   45
               TabIndex        =   19
               Top             =   525
               Width           =   1395
            End
            Begin VB.CheckBox chkProtesto 
               Alignment       =   1  'Right Justify
               Caption         =   "Protesto ?"
               Height          =   285
               Left            =   45
               TabIndex        =   18
               Top             =   210
               Width           =   1395
            End
            Begin VB.CheckBox chkCargoAuto 
               Alignment       =   1  'Right Justify
               Caption         =   "Cargo Automat."
               Height          =   285
               Left            =   45
               TabIndex        =   17
               Top             =   840
               Width           =   1395
            End
         End
         Begin VB.Frame fraDescPlan 
            Height          =   1005
            Left            =   5760
            TabIndex        =   10
            Top             =   1920
            Width           =   5280
            Begin VB.Label lblmodular 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1155
               TabIndex        =   15
               Top             =   585
               Width           =   2505
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cod Modular"
               Height          =   195
               Index           =   45
               Left            =   135
               TabIndex        =   14
               Top             =   630
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Crédito Personal Descuentos Por Planillas"
               Height          =   195
               Index           =   44
               Left            =   120
               TabIndex        =   13
               Top             =   0
               Width           =   2955
            End
            Begin VB.Label lblinstitucion 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1155
               TabIndex        =   12
               Top             =   255
               Width           =   3960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Institución"
               Height          =   195
               Index           =   43
               Left            =   135
               TabIndex        =   11
               Top             =   300
               Width           =   720
            End
         End
         Begin VB.Frame fraCreditosRefinanciado 
            Caption         =   "Creditos Refinaciados"
            Height          =   1290
            Left            =   8040
            TabIndex        =   8
            Top             =   240
            Width           =   3045
            Begin MSComctlLib.ListView lstRefinanciados 
               Height          =   1080
               Left            =   60
               TabIndex        =   9
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
               NumItems        =   4
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
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Otros"
                  Object.Width           =   1499
               EndProperty
            End
         End
         Begin VB.Label Label16 
            Caption         =   "Glosa:"
            Height          =   255
            Left            =   120
            TabIndex        =   165
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblFecEEFF 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3480
            TabIndex        =   145
            Top             =   1950
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ult. Fecha EE.FF."
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
            Index           =   29
            Left            =   3480
            TabIndex        =   144
            Top             =   1680
            Width           =   1530
         End
         Begin VB.Label lblmetodoLiquidacion 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4800
            TabIndex        =   30
            Top             =   1320
            Width           =   660
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Método Liquidación :"
            Height          =   195
            Index           =   47
            Left            =   3240
            TabIndex        =   29
            Top             =   1365
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
            TabIndex        =   28
            Top             =   1695
            Width           =   660
         End
         Begin VB.Label lblfechajudicial 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1920
            TabIndex        =   27
            Top             =   1965
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de  Ing. Judicial :"
            Height          =   195
            Index           =   41
            Left            =   120
            TabIndex        =   26
            Top             =   1965
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
            TabIndex        =   25
            Top             =   1065
            Width           =   2250
         End
         Begin VB.Label lblfechaCancelacion 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1920
            TabIndex        =   24
            Top             =   1320
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Cancelación :"
            Height          =   195
            Index           =   39
            Left            =   120
            TabIndex        =   23
            Top             =   1335
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
            TabIndex        =   22
            Top             =   165
            Width           =   1890
         End
         Begin VB.Label lblrechazo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   720
            TabIndex        =   21
            Top             =   405
            Width           =   4065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Motivo :"
            Height          =   195
            Index           =   37
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   570
         End
      End
      Begin VB.Frame fragarantias 
         Height          =   2895
         Left            =   -74400
         TabIndex        =   5
         Top             =   510
         Width           =   10335
         Begin MSComctlLib.ListView lstgarantias 
            Height          =   2475
            Left            =   360
            TabIndex        =   6
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
            NumItems        =   7
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
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Garantia"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.Label lblLiquidacion 
         Caption         =   "Crédito requiere actualización de liquidación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   240
         TabIndex        =   173
         Top             =   3600
         Visible         =   0   'False
         Width           =   6795
      End
      Begin VB.Label lblMensajeMIVIVIENDA 
         Caption         =   "Los Montos de cancelación son referenciales ya que no se está considerando la liquidación del Bono y/o Premio al Buen Pagador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   7200
         TabIndex        =   171
         Top             =   3480
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label lbl_Reprogramado 
         Caption         =   "REPROGRAMADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   -74400
         TabIndex        =   170
         Top             =   3360
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label lblMonedaP 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "SOLES"
         Height          =   240
         Left            =   8700
         TabIndex        =   140
         Top             =   3240
         Width           =   1635
      End
      Begin VB.Label Label8 
         Caption         =   "Moneda:"
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
         Left            =   7845
         TabIndex        =   139
         Top             =   3240
         Width           =   780
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de  Ing. Judicial :"
      Height          =   195
      Index           =   27
      Left            =   480
      TabIndex        =   143
      Top             =   270
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1920
      TabIndex        =   142
      Top             =   270
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
      Index           =   14
      Left            =   480
      TabIndex        =   141
      Top             =   0
      Width           =   660
   End
End
Attribute VB_Name = "frmCredConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objPista As COMManejador.Pista 'MAVM 20100726
Public bVB As Boolean '*****RECO 20130701***********
Dim fbEsCredMIVIVIENDA As Boolean 'WIOR 20151220
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018

Public Sub ConsultaCliente(ByVal psCtaCod As String)
    ActxCta.NroCuenta = psCtaCod
    MostrarFecCompromisoPago 'JACA 20120117 para visualizar la pestaña de Compromiso de Pago
    Call ActxCta_KeyPress(13)
    CmdNuevaCons.Enabled = False
    Me.Show 1
End Sub

'Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
'
'Dim oCredPersRela As COMDCredito.UCOMCredRela
'Dim L As ListItem
'Dim RDatos As ADODB.Recordset
'Dim R As ADODB.Recordset
'Dim oCred As COMDCredito.DCOMCredito
'Dim oCalend As COMDCredito.DCOMCalendario
'Dim oMontoDesemb As Double
'Dim nCapPag As Double
'Dim nIntPag As Double
'Dim nGasto As Double
'Dim nitf As Double
'Dim nMora As Double
'Dim oNegCred As COMNCredito.NCOMCredito
'Dim MatCalend As Variant
'Dim MatCalendPar As Variant
'Dim oGarantia As COMDCredito.DCOMGarantia
'Dim i As Integer
'Dim MatGastosCancelacion As Variant
'Dim nNumGastosCancel As Integer
'Dim oGastos As COMNCredito.NCOMGasto
'Dim oNegCredito As COMNCredito.NCOMCredito
'Dim MatCalendDistribuido As Variant
'Dim nDiasAtraso As Integer
'Dim nPrdEstadoTmp As Long
'Dim oDCredDoc As COMDCredito.DCOMCredDoc
'
'    On Error GoTo ErrorCargaDatos
'
'    Screen.MousePointer = 11
'    CargaDatos = True
'    'Carga Relaciones de Credito
'    Set oCredPersRela = New COMDCredito.UCOMCredRela
'    Call oCredPersRela.CargaRelacPersCred(psCtaCod)
'    oCredPersRela.IniciarMatriz
'    listaClientes.ListItems.Clear
'    Do While Not oCredPersRela.EOF
'        Set L = listaClientes.ListItems.Add(, , oCredPersRela.ObtenerNombre)
'        L.SubItems(1) = oCredPersRela.ObtenerRelac
'        L.SubItems(2) = oCredPersRela.ObtenerCodigo
'        oCredPersRela.siguiente
'    Loop
'    Set oCredPersRela = Nothing
'
'    'Carga Datos del Credito
'    Set oCred = New COMDCredito.DCOMCredito
'    Set RDatos = oCred.RecuperaConsultaCred(psCtaCod)
'    Set oCred = Nothing
'    If RDatos.BOF And RDatos.EOF Then
'        CargaDatos = False
'        'Call CargaDatos(psCtaCod)
'        RDatos.Close
'        Set RDatos = Nothing
'        Screen.MousePointer = 0
'        Exit Function
'    End If
'
'    LblIntComVen.Caption = Format(RDatos!nTasaCompVen, "#0.0000")
'    LblIntCom.Caption = Format(RDatos!nTasaComp, "#0.0000")
'    LblIntMor.Caption = Format(RDatos!nTasaMor, "#0.0000")
'    ChkCalDin.value = RDatos!nCalendDinamico
'    ChkCuotaCom.value = RDatos!bCuotaCom
'    ChkMiViv.value = RDatos!bMiVivienda
'    nDiasAtraso = IIf(IsNull(RDatos!nDiasAtraso), 0, RDatos!nDiasAtraso)
'    lbltipoCredito.Caption = Trim(RDatos!cTipoCredDescrip)
'    lblEstado.Caption = Trim(RDatos!cEstActual)
'    nPrdEstadoTmp = RDatos!nPrdEstado
'    lblfuente.Caption = Trim(IIf(IsNull(RDatos!cFteIngreso), "", RDatos!cFteIngreso))
'    lblLinea.Caption = Trim(IIf(IsNull(RDatos!cLineaDesc), "", RDatos!cLineaDesc))
'    lblanalista.Caption = Trim(IIf(IsNull(RDatos!cAnalista), "", RDatos!cAnalista))
'    lblapoderado.Caption = Trim(IIf(IsNull(RDatos!cApoderado), "", RDatos!cApoderado))
'    lblcondicion.Caption = Trim(IIf(IsNull(RDatos!cCondicion), "", RDatos!cCondicion))
'    lbldestino.Caption = Trim(IIf(IsNull(RDatos!cDestino), "", RDatos!cDestino))
'    lblnota1.Caption = IIf(IsNull(RDatos!nNota), "", RDatos!nNota)
'    lbltasainteres.Caption = Format(IIf(IsNull(RDatos!nTasaInteres), 0, RDatos!nTasaInteres), "#0.00")
'    lbltipocuota.Caption = Trim(IIf(IsNull(RDatos!cTipoCuota), "", RDatos!cTipoCuota))
'    If IsNull(RDatos!dvigencia) Then
'        lblfechavigencia.Caption = ""
'    Else
'        lblfechavigencia.Caption = Format(RDatos!dvigencia, "dd/mm/yyyy")
'    End If
'
'    'Ficha de Historial
'    If IsNull(RDatos!dFecSol) Then
'        lblfechsolicitud.Caption = ""
'    Else
'        lblfechsolicitud.Caption = Format(RDatos!dFecSol, "dd/mm/yyyy")
'    End If
'    lblMontoSolicitado.Caption = Format(IIf(IsNull(RDatos!nMontoSol), 0, RDatos!nMontoSol), "#0.00")
'    lblcuotasSolicitud.Caption = IIf(IsNull(RDatos!nCuotasSol), "", RDatos!nCuotasSol)
'    lblPlazoSolicitud.Caption = IIf(IsNull(RDatos!nPlazoSol), 0, RDatos!nPlazoSol)
'    If IsNull(RDatos!dFecSug) Then
'        lblfechasugerida.Caption = ""
'    Else
'        lblfechasugerida.Caption = Format(RDatos!dFecSug, "dd/mm/yyyy")
'    End If
'    lblMontosugerido.Caption = Format(IIf(IsNull(RDatos!nMontoSug), 0, RDatos!nMontoSug), "#0.00")
'    lblcuotasugerida.Caption = Format(IIf(IsNull(RDatos!nCuotasSug), 0, RDatos!nCuotasSug), "#0")
'    lblPlazoSugerido.Caption = IIf(IsNull(RDatos!nPlazoSug), 0, RDatos!nPlazoSug)
'    lblmontoCuotaSugerida.Caption = Format(IIf(IsNull(RDatos!nCuotaSug), 0, RDatos!nCuotaSug), "#0.00")
'    lblGraciasugerida.Caption = IIf(IsNull(RDatos!nPeriodoGraciaSug), 0, RDatos!nPeriodoGraciaSug)
'    If IsNull(RDatos!dFecApr) Then
'        lblfechaAprobado.Caption = ""
'    Else
'        lblfechaAprobado.Caption = Format(RDatos!dFecApr, "dd/mm/yyyy")
'    End If
'    lblMontoAprobado.Caption = Format(IIf(IsNull(RDatos!nMontoApr), 0, RDatos!nMontoApr), "#0.00")
'    lblcuotasAprobado.Caption = IIf(IsNull(RDatos!nCuotasApr), 0, RDatos!nCuotasApr)
'    lblPlazoAprobado.Caption = IIf(IsNull(RDatos!nPlazoApr), 0, RDatos!nPlazoApr)
'    lblmontoCuotaAprobada.Caption = Format(IIf(IsNull(RDatos!nCuotaApr), 0, RDatos!nCuotaApr), "#0.00")
'    lblGraciaAprobada.Caption = IIf(IsNull(RDatos!nPeriodoGraciaApr), 0, RDatos!nPeriodoGraciaApr)
'    lblIntGraciaApr.Caption = Format(IIf(IsNull(RDatos!nTasaGracia), 0, RDatos!nTasaGracia), "#0.00")
'    lblTipoGraciaApr.Caption = IIf(IsNull(RDatos!cTipoGracia), "", RDatos!cTipoGracia)
'
'    'Ficha Desembolsos Realizados
'    listaDesembolsos.ListItems.Clear
'    Set oCalend = New COMDCredito.DCOMCalendario
'    Set R = oCalend.RecuperaCalendarioDesemb(psCtaCod)
'    Set oCalend = Nothing
'    oMontoDesemb = 0
'    Do While Not R.EOF
'        Set L = listaDesembolsos.ListItems.Add(, , Format(R!dPago, "dd/mm/yyyy"))
'        L.SubItems(1) = R!nCuota
'        L.SubItems(2) = Format(R!nCapital, "#0.00")
'        L.SubItems(3) = Format(R!nGasto, "#0.00")
'        L.SubItems(4) = IIf(R!nColocCalendEstado = gColocCalendEstadoPendiente, "PENDIENTE", "DESEMBOLSADO")
'        If R!nColocCalendEstado = gColocCalendEstadoPagado Then
'            oMontoDesemb = oMontoDesemb + R!nCapital
'        End If
'        R.MoveNext
'    Loop
'    lbltipoDesembolso.Caption = IIf(IsNull(RDatos!cTipoDesemb), "", RDatos!cTipoDesemb)
'    lblmontoDesembolsado.Caption = Format(oMontoDesemb, "#0.00")
'    lbltotalDesembolso.Caption = Format(IIf(IsNull(RDatos!nMontoApr), 0, RDatos!nMontoApr), "#0.00")
'    R.Close
'    Set R = Nothing
'
'    'Ficha Pagos Realizados
'    'Set oCalend = New Dcalendario
'    'Set R = oCalend.RecuperaCalendarioPagosRealizados(psCtaCod)
'    Set oCred = New COMDCredito.DCOMCredito
'    'Set R = oCred.RecuperaPagosRealizados(psCtaCod)
'     Set R = oCred.RecuperaDetallePago(psCtaCod)
'    Set oCred = Nothing
'    ListaPagos.ListItems.Clear
'    nCapPag = 0
'    nIntPag = 0
'    nGasto = 0
'    nMora = 0
'    nitf = 0
'
'    Do While Not R.EOF
'        Set L = ListaPagos.ListItems.Add(, , Format(R!dFecPago, "dd/mm/yyyy"))
'        L.SubItems(1) = R!nNroCuota
'        L.SubItems(2) = Format(R!nMontoPagado, "#0.00")
'        L.SubItems(3) = Format(R!nCapital, "#0.00")
'        nCapPag = nCapPag + CDbl(L.SubItems(3))
'        L.SubItems(4) = Format(R!nInteres, "#0.00")
'        nIntPag = nIntPag + CDbl(L.SubItems(4))
'        L.SubItems(5) = Format(R!nMora, "#0.00")
'        nMora = nMora + CDbl(L.SubItems(5))
'        L.SubItems(6) = Format(R!nGastos, "#0.00")
'        nGasto = nGasto + CDbl(L.SubItems(6))
'
'        L.SubItems(7) = Format(R!nitf, "#0.00")
'        nitf = nitf + CDbl(L.SubItems(7))
'
'
'        L.SubItems(8) = Format(R!nDiasMora, "#0")
'        L.SubItems(9) = Format(R!nSaldoCap, "#0.00")
'        L.SubItems(10) = IIf(IsNull(R!cusuario), "", R!cusuario)
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'
'    lblcapitalpagado.Caption = Format(nCapPag, "#0.00")
'    lblintcompPag.Caption = Format(nIntPag, "#0.00")
'    lblIntMorPag.Caption = Format(nMora, "#0.00")
'    lblGastopagado.Caption = Format(nGasto, "#0.00")
'    lblITFPagado.Caption = Format(nitf, "#0.00")
'
'    'Ficha Pagos Pendientes
'    Set oNegCred = New COMNCredito.NCOMCredito
'    MatCalend = oNegCred.RecuperaMatrizCalendarioPendiente(psCtaCod)
'    MatCalendPar = oNegCred.RecuperaMatrizCalendarioPendiente(psCtaCod, True)
'    MatCalend = UnirMatricesMiViviendaAmortizacion(MatCalend, MatCalendPar)
'    lstCuotasPend.ListItems.Clear
'    For i = 0 To UBound(MatCalend) - 1
'        Set L = lstCuotasPend.ListItems.Add(, , MatCalend(i, 1))
'        L.SubItems(1) = MatCalend(i, 0)
'        L.SubItems(2) = MatCalend(i, 3)
'        L.SubItems(3) = Format(CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 11)), "#0.00")
'        L.SubItems(4) = MatCalend(i, 6)
'        L.SubItems(8) = DateDiff("d", CDate(MatCalend(i, 0)), gdFecSis)
'        L.SubItems(5) = MatCalend(i, 9)
'        L.SubItems(6) = fgITFCalculaImpuesto(CDbl(Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 11)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 9)), "#0.00")))
'        L.SubItems(7) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 11)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 9)) + CDbl(L.SubItems(6)), "#0.00")
'        L.SubItems(9) = MatCalend(i, 10)
'    Next i
'
'    lblCapitalCuoPend.Caption = Format(oNegCred.MatrizCapitalVencido(MatCalend, gdFecSis), "#0.00")
'    lblInteresCuoPend.Caption = Format(oNegCred.MatrizIntCompVencido(MatCalend, gdFecSis) + _
'                                oNegCred.MatrizIntCompVencVencido(MatCalend, gdFecSis) + _
'                                oNegCred.MatrizIntGraciaVencido(MatCalend, gdFecSis) + _
'                                oNegCred.MatrizIntReprogVencido(MatCalend, gdFecSis) + _
'                                oNegCred.MatrizIntSuspensoVencido(MatCalend, gdFecSis), "#0.00")
'    'lblMoraCuoPend.Caption = Format(oNegCred.MatrizInteresMorFecha(psCtaCod, MatCalend), "#0.00")
'    'lblGastoCuoPend.Caption = Format(oNegCred.MatrizGastosVencidos(MatCalend, gdFecSis), "#0.00")
'    lblMoraCuoPend.Caption = Format(ObtenerMoraVencida(gdFecSis, MatCalend), "#0.00")
'    lblGastoCuoPend.Caption = Format(ObtenerGastoVencido(gdFecSis, MatCalend), "#0.00")
'    lblTotalCuoPend.Caption = Format(CDbl(lblCapitalCuoPend.Caption) + CDbl(lblInteresCuoPend.Caption) + CDbl(lblMoraCuoPend.Caption) + CDbl(lblGastoCuoPend.Caption), "#0.00")
'
'    lblSaldoKCalend.Caption = oNegCred.MatrizCapitalCalendario(MatCalend)
'    lblIntCompCalend.Caption = Format(oNegCred.MatrizIntCompCalendario(MatCalend) + _
'                            oNegCred.MatrizIntComVencCalendario(MatCalend) + _
'                            oNegCred.MatrizIntGraciaCalendario(MatCalend) + _
'                            oNegCred.MatrizIntReprogCalendario(MatCalend) + _
'                            oNegCred.MatrizIntSuspensoCalendario(MatCalend), "#0.00")
'    lblGastoCalend.Caption = Format(oNegCred.MatrizIntGastosCalendario(MatCalend), "#0.00")
'    lblIntMorCalend.Caption = Format(oNegCred.MatrizIntMoratorioCalendario(MatCalend), "#0.00")
'    lblTotalCalend.Caption = Format(CDbl(lblSaldoKCalend.Caption) + CDbl(lblIntCompCalend.Caption) + CDbl(lblGastoCalend.Caption) + CDbl(lblIntMorCalend.Caption), "#0.00")
'
'    lblSaldoKFecha.Caption = Format(oNegCred.MatrizCapitalAFecha(psCtaCod, MatCalend), "#0.00")
'    lblGastoFecha.Caption = 0#
'    If UBound(MatCalend) > 0 Then
'
'        lblIntCompFecha.Caption = Format(oNegCred.MatrizInteresTotalesAFechaSinMora(psCtaCod, MatCalend, gdFecSis), "#0.00")
'        'lblGastoFecha.Caption = Format(oNegCred.MatrizGastosFecha(psCtaCod, MatCalend), "#0.00")
'        'lblIntMorFecha.Caption = Format(oNegCred.MatrizInteresMorFecha(psCtaCod, MatCalend), "#0.00")
'        lblIntMorFecha.Caption = Format(ObtenerMoraVencida(gdFecSis, MatCalend), "#0.00")
'        lblGastoFecha.Caption = Format(ObtenerGastoVencido(gdFecSis, MatCalend), "#0.00")
'        'lblPenalidadFecha.Caption = Format(oNegCred.CalculaGastoPenalidadCancelacion(CDbl(lblSaldoKFecha.Caption), CInt(Mid(psCtaCod, 9, 1))), "#0.00")
'        lblPenalidadFecha.Caption = "0.00"
'        lblTotalFecha.Caption = Format(CDbl(lblSaldoKFecha.Caption) + CDbl(lblIntCompFecha.Caption) + CDbl(lblGastoFecha.Caption) + CDbl(lblIntMorFecha.Caption) + CDbl(lblPenalidadFecha.Caption), "#0.00")
'        Set oNegCredito = New COMNCredito.NCOMCredito
'        If nPrdEstadoTmp = 2020 Or nPrdEstadoTmp = 2021 Or nPrdEstadoTmp = 2022 Or nPrdEstadoTmp = 2030 Or nPrdEstadoTmp = 2031 Or nPrdEstadoTmp = 2032 Then
'            Set oGastos = New COMNCredito.NCOMGasto
'            MatGastosCancelacion = oGastos.GeneraCalendarioGastos(Array(0), Array(0), nNumGastosCancel, gdFecSis, ActxCta.NroCuenta, 1, "CA", , , CDbl(lblTotalFecha.Caption), oNegCredito.MatrizMontoCapitalAPagar(MatCalend, gdFecSis), oNegCredito.MatrizCuotaPendiente(MatCalend, MatCalend), , , , , nDiasAtraso)
'            lblGastoFecha.Caption = Format(CDbl(lblGastoFecha.Caption) + Format(MontoTotalGastosGenerado(MatGastosCancelacion, nNumGastosCancel, Array("CA", "PA", "")), "#0.00"), "#0.00")
'            lblTotalFecha.Caption = Format(CDbl(lblTotalFecha.Caption) + CDbl(Format(MontoTotalGastosGenerado(MatGastosCancelacion, nNumGastosCancel, Array("CA", "PA", "")), "#0.00")), "#0.00")
'            Set oGastos = Nothing
'            Set oNegCredito = Nothing
'        End If
'    End If
'    If UBound(MatCalend) > 0 Then
'        If CDate(MatCalend(0, 0)) < gdFecSis Then
'            'lblGastoCuoPend.Caption = lblGastoFecha.Caption
'            lblTotalCuoPend.Caption = Format(CDbl(lblCapitalCuoPend.Caption) + CDbl(lblInteresCuoPend.Caption) + CDbl(lblMoraCuoPend.Caption) + CDbl(lblGastoCuoPend.Caption), "#0.00")
'        End If
'    End If
'    'Ficha de Garantias
'    lstgarantias.ListItems.Clear
'    Set oGarantia = New COMDCredito.DCOMGarantia
'    Set R = oGarantia.RecuperaGarantiaCredito(psCtaCod)
'    Set oGarantia = Nothing
'    Do While Not R.EOF
'        Set L = lstgarantias.ListItems.Add(, , Trim(R!cTpoGarantia))
'        L.SubItems(1) = Trim(R!cDescripcion)
'        L.SubItems(2) = Trim(R!cDocDesc)
'        L.SubItems(3) = Trim(R!cNroDoc)
'        L.SubItems(4) = Trim(R!cMoneda)
'        'L.SubItems(5) = Format(R!nGravado, "#0.00")
'        L.SubItems(5) = Format(R!nRealizacion, "#0.00")
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'
'    'Ficha de Otros Datos
'    If IsNull(RDatos!cMotivoRech) Then
'        lblrechazo.Caption = ""
'    Else
'        lblrechazo.Caption = Trim(RDatos!cMotivoRech)
'    End If
'    If IsNull(RDatos!dFecCancel) Then
'        lblfechaCancelacion.Caption = ""
'    Else
'        lblfechaCancelacion.Caption = Format(RDatos!dFecCancel, "dd/mm/yyyy")
'    End If
'    lblmetodoLiquidacion.Caption = Trim(IIf(IsNull(RDatos!cMetLiquidacion), "", RDatos!cMetLiquidacion))
'    If IsNull(RDatos!dFecJud) Then
'        lblfechajudicial.Caption = ""
'    Else
'        lblfechajudicial.Caption = Format(RDatos!dFecJud, "dd/mm/yyyy")
'    End If
'
'    If IsNull(RDatos!cProtesto) Then
'        chkProtesto.value = 0
'    Else
'        If Trim(RDatos!cProtesto) = "1" Then
'            chkProtesto.value = 1
'        Else
'            chkProtesto.value = 0
'        End If
'    End If
'    If IsNull(RDatos!nEstRefin) Then
'        chkRefinanciado.value = 0
'    Else
'        If RDatos!nEstRefin = 2030 Or RDatos!nEstRefin = 2031 Or RDatos!nEstRefin = 2032 Then
'            chkRefinanciado.value = 1
'        Else
'            chkRefinanciado.value = 0
'        End If
'    End If
'
'    If IsNull(RDatos!bCargoAuto) Then
'        chkCargoAuto.value = 0
'    Else
'        If RDatos!bCargoAuto = True Then
'            chkCargoAuto.value = 1
'        Else
'            chkCargoAuto.value = 0
'        End If
'    End If
'
'    lstRefinanciados.ListItems.Clear
'    Set oCred = New COMDCredito.DCOMCredito
'    Set R = oCred.RecuperaCreditosRefinanciados(psCtaCod)
'    Set oCred = Nothing
'    Do While Not R.EOF
'        Set L = lstRefinanciados.ListItems.Add(, , R!cCtaCodRef)
'        L.SubItems(1) = Format(IIf(IsNull(R!nCapitalRef), 0, R!nCapitalRef), "#0.00")
'        L.SubItems(2) = Format(IIf(IsNull(R!nInteresRef), 0, R!nInteresRef), "#0.00")
'       ' L.SubItems(3) = Format(IIf(IsNull(R!nGastosRef), 0, R!nGastosRef), "#0.00")
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'
'    If CInt(Mid(psCtaCod, 6, 3)) = gColConsuDctoPlan Then
'        lblinstitucion.Caption = PstaNombre(RDatos!cConvenio)
'        lblmodular.Caption = Trim(RDatos!cCodModular)
'    End If
'    RDatos.Close
'    Set RDatos = Nothing
'    Set oNegCred = Nothing
'    'RecupCreditoAntiguo ActxCta.NroCuenta
'    'Recup_FechaCancelacion ActxCta.NroCuenta
'
'    Set oDCredDoc = New COMDCredito.DCOMCredDoc
'    Dim sCtaCodAnt As String
'    Dim sFecha As String
'
'    sCtaCodAnt = oDCredDoc.Recup_CreditoAntiguo(psCtaCod)
'    sFecha = oDCredDoc.Recup_FechaCancelacion(psCtaCod)
'
'    Set oDCredDoc = Nothing
'
'    If Len(sCtaCodAnt) = 0 Then
'        LblAntigua = "ES UN CREDITO NUEVO"
'    Else
'        LblAntigua.Caption = sCtaCodAnt
'    End If
'
'    txtFechaCancelacion.Text = sFecha
'
'    Screen.MousePointer = 0
'    Exit Function
'
'ErrorCargaDatos:
'        MsgBox Err.Description, vbCritical, "Aviso"
'
'End Function

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not CargaDatos(ActxCta.NroCuenta) Then
            ActxCta.NroCuenta = ""
            ActxCta.Enabled = True
            LimpiaControles Me, True
            MsgBox "No se encontro el Credito", vbInformation, "Aviso"
            ActxCta.NroCuenta = ""
            ActxCta.CMAC = gsCodCMAC
            ActxCta.Age = gsCodAge
        Else
            ActxCta.Enabled = False
        End If
    End If
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean

Dim L As MSComctlLib.ListItem
Dim RDatos As ADODB.Recordset
Dim nMontoDesemb As Double

Dim nMora As Double
Dim oCred As COMNCredito.NCOMCredito
Dim oLeasing As COMNCredito.NCOMLeasing 'ORCR 20140414
Dim MatCalend As Variant
Dim i As Integer
Dim nCapPag As Double
Dim nIntPag As Double
Dim nGasto As Double
Dim nITF As Double

'Variables para el uso de los COMPONENTES

Dim MatRelac As Variant
Dim rsCalen As ADODB.Recordset
Dim rsPag As ADODB.Recordset
Dim MatImpuesto As Variant
Dim sCapitalCuoPend As String
Dim sInteresCuoPend As String
Dim sMoraCuoPend As String
Dim sGastoCuoPend As String
Dim sTotalCuoPend As String
Dim sSaldoKCalend As String
Dim sIntCompCalend As String
Dim sGastoCalend As String
Dim sIntMorCalend As String
Dim sTotalCalend As String
Dim sSaldoKFecha As String
Dim sIntCompFecha As String
Dim sIntMorFecha As String
Dim sGastoFecha As String
Dim sPenalidadFecha As String
Dim sTotalFecha As String
Dim rsGar As ADODB.Recordset
Dim rsRef As ADODB.Recordset
Dim sCtaCodAnt As String
Dim sFecha As String
fbEsCredMIVIVIENDA = False 'WIOR 20151224
Dim nIntCompVencido As Double 'RIRO 20210430
Dim nIntCompVencidoPag As Double 'RIRO 20210430
Dim nIntCompVencidoPend As Double 'RIRO 20210430
'----------------------------------------
    On Error GoTo ErrorCargaDatos
    
    Screen.MousePointer = 11
    CargaDatos = True
    
    Set oCred = New COMNCredito.NCOMCredito
    nIntCompVencido = 0
    CargaDatos = oCred.CargaDatosHistorialCredito(psCtaCod, gdFecSis, MatRelac, RDatos, rsCalen, rsPag, MatCalend, _
                                                MatImpuesto, sCapitalCuoPend, sInteresCuoPend, sMoraCuoPend, _
                                                sGastoCuoPend, sTotalCuoPend, sSaldoKCalend, sIntCompCalend, _
                                                sGastoCalend, sIntMorCalend, sTotalCalend, sSaldoKFecha, _
                                                sIntCompFecha, sIntMorFecha, sGastoFecha, sPenalidadFecha, sTotalFecha, _
                                                rsGar, rsRef, sCtaCodAnt, sFecha, nIntCompVencido)
    Set oCred = Nothing
    'MAVM 20100726 ***
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gConsultar, "Consulta Historial de Credito", psCtaCod, gCodigoCuenta
    Set objPista = Nothing
    '***
    
    'Carga Relaciones de Credito
    listaClientes.ListItems.Clear
    For i = 1 To UBound(MatRelac)
        Set L = listaClientes.ListItems.Add(, , MatRelac(i, 1))
        L.SubItems(1) = MatRelac(i, 2)
        L.SubItems(2) = MatRelac(i, 3)
    Next i
    lblMensajeMIVIVIENDA.Visible = False 'WIOR 20151224
    If RDatos.BOF And RDatos.EOF Then
        CargaDatos = False
        'Call CargaDatos(psCtaCod)
        'RDatos.Close
        'Set RDatos = Nothing
        Screen.MousePointer = 0
        Exit Function
    End If

'Segmentacion de riesgo JOEP20201118
    If Not (RDatos.BOF And RDatos.EOF) Then
        If RDatos!cNivelRiesgoCredito <> "" Then
            txtFechaCierre.Text = RDatos!dFechaCierreUltimo
            txtNivel.Text = RDatos!cNivelRiesgoCredito
            txtMotivo.Text = RDatos!NivelRiesgoCalculo
        End If
    End If
'Segmentacion de riesgo JOEP20201118

    LblIntComVen.Caption = Format(RDatos!nTasaCompVen, "#0.0000")
    LblIntCom.Caption = Format(RDatos!nTasaComp, "#0.0000")
    LblIntMor.Caption = Format(RDatos!nTasaMor, "#0.0000")
    ChkCalDin.value = RDatos!nCalendDinamico
    ChkCuotaCom.value = RDatos!bCuotaCom
    ChkMiViv.value = RDatos!bMiVivienda
    lbltipoCredito.Caption = Trim(RDatos!cTipoCredDescrip)
    lblTipoProducto.Caption = Trim(RDatos!cTipoProdDescrip)
    LblEstado.Caption = Trim(RDatos!cEstActual)
    lblfuente.Caption = Trim(IIf(IsNull(RDatos!cFteIngreso), "", RDatos!cFteIngreso))
    lblLinea.Caption = Trim(IIf(IsNull(RDatos!cLineaDesc), "", RDatos!cLineaDesc))
    lblanalista.Caption = Trim(IIf(IsNull(RDatos!cAnalista), "", RDatos!cAnalista))
    lblapoderado.Caption = Trim(IIf(IsNull(RDatos!cApoderado), "", RDatos!cApoderado))
    lblcondicion.Caption = Trim(IIf(IsNull(RDatos!cCondicion), "", RDatos!cCondicion))
    lbldestino.Caption = Trim(IIf(IsNull(RDatos!cDestino), "", RDatos!cDestino))
    lblnota1.Caption = IIf(IsNull(RDatos!nNota), "", RDatos!nNota)
    'lbltasainteres.Caption = Format(IIf(IsNull(RDatos!nTasaInteres), 0, RDatos!nTasaInteres), "#0.00")
    lbltasainteres.Caption = Format(IIf(IsNull(RDatos!nTasaInteres), 0, RDatos!nTasaInteres), "#0.0000") 'JUEZ 20140221 Para redondear a 4 decimales
    lbltipocuota.Caption = Trim(IIf(IsNull(RDatos!cTipoCuota), "", RDatos!cTipoCuota))
    lblHonrado.Caption = RDatos!cHonrado 'APRI20210327 POR HONRAMIENTO
    If IsNull(RDatos!dVigencia) Then
        lblfechavigencia.Caption = ""
    Else
        lblfechavigencia.Caption = Format(RDatos!dVigencia, "dd/mm/yyyy")
    End If
    'WIOR 20151224 ***
     Set oCred = New COMNCredito.NCOMCredito
    fbEsCredMIVIVIENDA = oCred.EsCredMIVIVENDA(RDatos!nTipoProdCod, RDatos!nTipoCredCod, 3)
    
    lblMensajeMIVIVIENDA.Visible = False
    If fbEsCredMIVIVIENDA Then
        lblMensajeMIVIVIENDA.Visible = True
    End If
    Set oCred = Nothing
    'WIOR FIN ********
    
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
    lblcuotasugerida.Caption = Format(IIf(IsNull(RDatos!nCuotasSug), 0, RDatos!nCuotasSug), "#0")
    lblPlazoSugerido.Caption = IIf(IsNull(RDatos!nPlazoSug), 0, RDatos!nPlazoSug)
    lblmontoCuotaSugerida.Caption = Format(IIf(IsNull(RDatos!nCuotaSug), 0, RDatos!nCuotaSug), "#0.00")
    lblGraciasugerida.Caption = IIf(IsNull(RDatos!nPeriodoGraciaSug), 0, RDatos!nPeriodoGraciaSug)
    If IsNull(RDatos!dFecApr) Then
        lblfechaAprobado.Caption = ""
    Else
        lblfechaAprobado.Caption = Format(RDatos!dFecApr, "dd/mm/yyyy")
    End If
    lblMontoAprobado.Caption = Format(IIf(IsNull(RDatos!nMontoApr), 0, RDatos!nMontoApr), "#0.00")
    'MAVM 20121219 ***
    lblcuotasAprobado.Caption = IIf(IsNull(RDatos!nCuotasApr), 0, RDatos!nCuotasApr)
    
    'If RDatos!cTipoGracia = "PRIMERA CUOTA" Then
    '    If IIf(IsNull(RDatos!nPeriodoGraciaApr), 0, RDatos!nPeriodoGraciaApr) > 0 Then 'WIOR 20130105
    '        lblcuotasAprobado.Caption = IIf(IsNull(RDatos!nCuotasApr), 0, RDatos!nCuotasApr + 1)
    '    'WIOR 20130105 *********************************************************************
    '    Else
    '        lblcuotasAprobado.Caption = IIf(IsNull(RDatos!nCuotasApr), 0, RDatos!nCuotasApr)
    '    End If
    '    'WIOR FIN **************************************************************************
    'Else
    '    lblcuotasAprobado.Caption = IIf(IsNull(RDatos!nCuotasApr), 0, RDatos!nCuotasApr)
    'End If
    '***
    lblPlazoAprobado.Caption = IIf(IsNull(RDatos!nPlazoApr), 0, RDatos!nPlazoApr)
    lblmontoCuotaAprobada.Caption = Format(IIf(IsNull(RDatos!nCuotaApr), 0, RDatos!nCuotaApr), "#0.00")
    lblGraciaAprobada.Caption = IIf(IsNull(RDatos!nPeriodoGraciaApr), 0, RDatos!nPeriodoGraciaApr)
    lblIntGraciaApr.Caption = Format(IIf(IsNull(RDatos!nTasaGracia), 0, RDatos!nTasaGracia), "#0.00")
    lblTipoGraciaApr.Caption = IIf(IsNull(RDatos!cTipoGracia), "", RDatos!cTipoGracia)
    
    'Ficha Desembolsos Realizados
    listaDesembolsos.ListItems.Clear
    nMontoDesemb = 0
    Do While Not rsCalen.EOF
        Set L = listaDesembolsos.ListItems.Add(, , Format(rsCalen!dPago, "dd/mm/yyyy"))
        L.SubItems(1) = rsCalen!nCuota
        L.SubItems(2) = Format(rsCalen!nCapital, "#0.00")
        L.SubItems(3) = Format(rsCalen!nGasto, "#0.00")
        L.SubItems(4) = IIf(rsCalen!nColocCalendEstado = gColocCalendEstadoPendiente, "PENDIENTE", "DESEMBOLSADO")
        If rsCalen!nColocCalendEstado = gColocCalendEstadoPagado Then
            nMontoDesemb = nMontoDesemb + rsCalen!nCapital
        End If
        rsCalen.MoveNext
    Loop
    lbltipoDesembolso.Caption = IIf(IsNull(RDatos!cTipoDesemb), "", RDatos!cTipoDesemb)
    lblmontoDesembolsado.Caption = Format(nMontoDesemb, "#0.00")
    lbltotalDesembolso.Caption = Format(IIf(IsNull(RDatos!nMontoApr), 0, RDatos!nMontoApr), "#0.00")
    
    ListaPagos.ListItems.Clear
    nCapPag = 0
    nIntPag = 0
    nGasto = 0
    nMora = 0
    nITF = 0
    nIntCompVencidoPag = 0 'RIRO20210524

    Do While Not rsPag.EOF
        Set L = ListaPagos.ListItems.Add(, , Format(rsPag!dFecPago, "dd/mm/yyyy"))
        L.SubItems(1) = IIf(IsNull(rsPag!nNroCuota), 0, rsPag!nNroCuota)
        L.SubItems(2) = Format(rsPag!nMontoPagado, "#0.00")
        L.SubItems(3) = Format(rsPag!nCapital, "#0.00")
        nCapPag = nCapPag + CDbl(L.SubItems(3))
        
        L.SubItems(4) = Format(rsPag!nInteres - rsPag!nIntCompVenc, "#0.00") 'RIRO 20210527 ADD nIntCompVenc
        nIntPag = nIntPag + CDbl(L.SubItems(4))
        
        L.SubItems(5) = Format(rsPag!nMora, "#0.00")
        nMora = nMora + CDbl(L.SubItems(5))
        
        'RIRO 20210502 ADD IntCompVencido ****
        L.SubItems(6) = Format(rsPag!nIntCompVenc, "#0.00")
        nIntCompVencidoPag = nIntCompVencidoPag + CDbl(L.SubItems(6))
        'END RIRO ****************************
        
        'L.SubItems(6) = Format(rsPag!nGastos, "#0.00")
        L.SubItems(7) = Format(rsPag!nGastoComision + rsPag!nGastoSegDes + rsPag!nGastoPolizaIncendio + rsPag!nGastoPolizaVehiculo, "#0.00") 'EJVG20140930
        nGasto = nGasto + CDbl(L.SubItems(7))
        
        'EJVG20140930 ***
        L.SubItems(8) = Format(rsPag!nGastoComision, "#0.00")
        L.SubItems(9) = Format(rsPag!nGastoSegDes, "#0.00")
        L.SubItems(10) = Format(rsPag!nGastoPolizaIncendio, "#0.00")
        L.SubItems(11) = Format(rsPag!nGastoPolizaVehiculo, "#0.00")
        'END EJVG *******
        L.SubItems(12) = Format(rsPag!nITF, "#0.00")
        nITF = nITF + CDbl(L.SubItems(12))
            
        L.SubItems(13) = Format(rsPag!nDiasMora, "#0")
        L.SubItems(14) = Format(rsPag!nSaldoCap, "#0.00")
        L.SubItems(15) = IIf(IsNull(rsPag!cUsuario), "", rsPag!cUsuario)
        rsPag.MoveNext
    Loop
    
    lblcapitalpagado.Caption = Format(nCapPag, "#0.00")
    lblintcompPag.Caption = Format(nIntPag, "#0.00")
    lblIntMorPag.Caption = Format(nMora, "#0.00")
    lblGastopagado.Caption = Format(nGasto, "#0.00")
    lblITFPagado.Caption = Format(nITF, "#0.00")
    lblIntCompVencido.Caption = Format(nIntCompVencidoPag, "#0.00")
    
    'Ficha Pagos Pendientes
    lstCuotasPend.ListItems.Clear
    For i = 0 To UBound(MatCalend) - 1
        Set L = lstCuotasPend.ListItems.Add(, , MatCalend(i, 1))
        L.SubItems(1) = MatCalend(i, 0)
        L.SubItems(2) = MatCalend(i, 3)
        'L.SubItems(3) = Format(CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 11)), "#0.00")
        L.SubItems(3) = Format(CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 8)), "#0.00")
        L.SubItems(4) = MatCalend(i, 11)
        L.SubItems(5) = MatCalend(i, 6)
        '**Modificado por DAOR 20080409 **********************************************
        'L.SubItems(8) = DateDiff("d", CDate(MatCalend(i, 0)), gdFecSis)
        If RDatos!nPrdEstado = gColocEstCancelado Or RDatos!nPrdEstado = gColocEstRefinanc Then
            L.SubItems(13) = DateDiff("d", CDate(MatCalend(i, 0)), RDatos!dPrdEstado)
        ElseIf RDatos!nPrdEstado = gColocEstRetirado Then '*** PEAC 20100608
            L.SubItems(13) = 0
        Else
        'ALPA 20110329***************************
            If (IIf(IsNull(RDatos!nDiasAtraso), 0, RDatos!nDiasAtraso) = 0 And i = 0) And DateDiff("d", RDatos!dVigencia, gdFecSis) <> 0 Then
                L.SubItems(13) = 0
            Else
                L.SubItems(13) = DateDiff("d", CDate(MatCalend(i, 0)), gdFecSis)
            End If
        End If
        '*****************************************************************************
        L.SubItems(6) = MatCalend(i, 9)
        'EJVG20140925 ***
        L.SubItems(7) = MatCalend(i, 13) 'Gasto Comisión(Todos los gastos no especificados)
        L.SubItems(8) = MatCalend(i, 14) 'Seg. Desgravamen
        L.SubItems(9) = MatCalend(i, 15) 'Seg. contra Incendio
        L.SubItems(10) = MatCalend(i, 16) 'Seg. Vehicular Multiriesgo
        L.SubItems(11) = MatImpuesto(i)
        L.SubItems(12) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 11)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 9)) + CDbl(MatImpuesto(i)), "#0.00")
        L.SubItems(14) = MatCalend(i, 10)
        'L.SubItems(6) = MatImpuesto(i)
        'L.SubItems(7) = Format(CDbl(MatCalend(i, 3)) + CDbl(MatCalend(i, 4)) + CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 6)) + CDbl(MatCalend(i, 7)) + CDbl(MatCalend(i, 11)) + CDbl(MatCalend(i, 8)) + CDbl(MatCalend(i, 9)) + CDbl(L.SubItems(6)), "#0.00")
        'L.SubItems(9) = MatCalend(i, 10)
        'END EJVG *******
    Next i
    
    lblCapitalCuoPend.Caption = sCapitalCuoPend
    lblInteresCuoPend.Caption = sInteresCuoPend
    lblMoraCuoPend.Caption = sMoraCuoPend
    lblGastoCuoPend.Caption = sGastoCuoPend
    lblTotalCuoPend.Caption = sTotalCuoPend
    lblIntCompVenc.Caption = Format(nIntCompVencido, "#0.00") 'RIRO 20210524
    
    lblSaldoKCalend.Caption = sSaldoKCalend
    lblIntCompCalend.Caption = sIntCompCalend
    lblGastoCalend.Caption = sGastoCalend
    lblIntMorCalend.Caption = sIntMorCalend
    lblTotalCalend.Caption = sTotalCalend
    lblSaldoKFecha.Caption = sSaldoKFecha
    lblGastoFecha.Caption = 0#
        
    'lblIntCompFecha.Caption = sIntCompFecha
    'lblIntMorFecha.Caption = sIntMorFecha
    'lblGastoFecha.Caption = sGastoFecha
    
    lblIntCompFecha.Caption = Format(IIf(IsNumeric(sIntCompFecha), (sIntCompFecha), 0), "#0.00")
    lblIntMorFecha.Caption = Format(IIf(IsNumeric(sIntMorFecha), (sIntMorFecha), 0), "#0.00")
    lblGastoFecha.Caption = Format(IIf(IsNumeric(sGastoFecha), (sGastoFecha), 0), "#0.00")
    
    'lblPenalidadFecha.Caption = sPenalidadFecha
    lblICV.Caption = Format(nIntCompVencido, "#0.00") 'RIRO 20210430
    lblTotalFecha.Caption = Format(IIf(IsNumeric(sTotalFecha), (sTotalFecha), 0), "#0.00")
    
    'Ficha de Garantias
    lstgarantias.ListItems.Clear
    Do While Not rsGar.EOF
        Set L = lstgarantias.ListItems.Add(, , Trim(rsGar!cTpoGarantia))
        L.SubItems(1) = Trim(rsGar!cDescripcion)
        L.SubItems(2) = Trim(rsGar!cDocDesc)
        L.SubItems(3) = Trim(rsGar!cNroDoc)
        L.SubItems(4) = Trim(rsGar!cMoneda)
        L.SubItems(5) = Format(rsGar!nRealizacion, "#0.00")
        L.SubItems(6) = rsGar!cNumGarantp 'EJVG20151014
        rsGar.MoveNext
    Loop
    
    'Ficha de Otros Datos
    If IsNull(RDatos!cMotivoRech) Then
        lblrechazo.Caption = ""
    Else
        lblrechazo.Caption = Trim(RDatos!cMotivoRech)
    End If
    'FRHU20130913 ***
    If IsNull(RDatos!cGlosa) Then
        'lblGlosa.Caption = ""
        txtGlosa.Text = ""
    Else
        'lblGlosa.Caption = Trim(RDatos!cGlosa)
        txtGlosa.Text = Trim(RDatos!cGlosa)
    End If
    'END FRHU *******
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
    
    
    '*** PEAC 20120814 - FICHA VISITAS DE GESTORES
    Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
    Set loPersCredito = New COMDColocRec.DCOMColRecCredito
    Dim lrVisit As ADODB.Recordset
    
'    If Len(psCtaCod) = 0 Then
'        MsgBox "Ingrese una cuenta", vbInformation + vbOKOnly, "Mensaje"
'        Exit Sub
'    End If
    
    If Len(psCtaCod) > 0 Then
    
        Set lrVisit = loPersCredito.ObtieneVisitaDeGestores(psCtaCod)
        Set loPersCredito = Nothing
        
        If Not lrVisit.EOF Then
'            MsgBox "No se encontraron datos.", vbInformation + vbOKOnly, "Mensaje"
'            Exit Sub
            Me.SSTab1.Tab = 1
            
            FeAdj.Clear
            FeAdj.FormaCabecera
            FeAdj.Rows = 2
            FeAdj.rsFlex = lrVisit
        End If
        
    End If
    '*** FIN PEAC
    
    
    'peac 20071228 eeff
    If IsNull(RDatos!dfecEEFF) Then
        lblFecEEFF.Caption = ""
    Else
        If Year(RDatos!dfecEEFF) <= "1950" Then
            lblFecEEFF.Caption = ""
        Else
            lblFecEEFF.Caption = Format(RDatos!dfecEEFF, "dd/mm/yyyy")
        End If
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
        If RDatos!nEstRefin = 2030 Or RDatos!nEstRefin = 2031 Or RDatos!nEstRefin = 2032 Then
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
    Do While Not rsRef.EOF
        Set L = lstRefinanciados.ListItems.Add(, , rsRef!cCtaCodRef)
        L.SubItems(1) = Format(IIf(IsNull(rsRef!nCapitalRef), 0, rsRef!nCapitalRef), "#0.00")
        L.SubItems(2) = Format(IIf(IsNull(rsRef!nInteresRef), 0, rsRef!nInteresRef), "#0.00")
        rsRef.MoveNext
    Loop
    
    If CInt(Mid(psCtaCod, 6, 3)) = gColConsuDctoPlan Then
        lblinstitucion.Caption = PstaNombre(IIf(IsNull(RDatos!cConvenio), "", RDatos!cConvenio))
        lblmodular.Caption = Trim(IIf(IsNull(RDatos!cCodModular), "", RDatos!cCodModular))
    End If
        
    If Len(sCtaCodAnt) = 0 Then
        LblAntigua = "ES UN CREDITO NUEVO"
    Else
        LblAntigua.Caption = sCtaCodAnt
    End If

    txtFechaCancelacion.Text = sFecha
          
    'INICIO ORCR-20140913*********
     Dim oCred2 As COMDCredito.DCOMCreditos
     Set oCred2 = New COMDCredito.DCOMCreditos
    
    lbl_Reprogramado.Visible = oCred2.CreditoReprogramado(psCtaCod)
    'FIN ORCR-20140913************
    
    'RIRO 20200911 ***********
    Dim oCred3 As COMNCredito.NCOMCredito
    Set oCred3 = New COMNCredito.NCOMCredito
    If Not oCred3.VerificaActualizacionLiquidacion(psCtaCod) Then
        lblLiquidacion.Visible = True
        lblTotalFecha.ForeColor = vbRed
    Else
        lblLiquidacion.Visible = False
        lblTotalFecha.ForeColor = vbBlack
    End If
    
    'lblLiquidacion
    
    'END RIRO ****************
    
    
    'ARCV 21-07-2006
    lblMonedaH.Caption = IIf(Mid(psCtaCod, 9, 1) = gMonedaNacional, "SOLES", "DOLARES")
    lblMonedaP.Caption = IIf(Mid(psCtaCod, 9, 1) = gMonedaNacional, "SOLES", "DOLARES")
    If Mid(psCtaCod, 9, 1) <> gMonedaNacional Then
        lblMonedaH.BackColor = vbGreen
        lblMonedaP.BackColor = vbGreen
    Else
        lblMonedaH.BackColor = vbWhite
        lblMonedaP.BackColor = vbWhite
    End If
    '---------------
       'ORCR INICIO 20140414 ***
    Dim sleasing As String
    sleasing = Mid(psCtaCod, 6, 3)
    '**ARLO20180712 ERS042 - 2018
    Set objProducto = New COMDCredito.DCOMCredito
    If objProducto.GetResultadoCondicionCatalogo("N0000092", sleasing) Then
    'If (sleasing = "515" Or sleasing = "516") Then
    '**ARLO20180712 ERS042 - 2018
        Set oLeasing = New COMNCredito.NCOMLeasing
        lbl_montoFinanciado.Caption = Format(oLeasing.ObtenerMontoFinanciado(psCtaCod), "#0.00")
        Call mostrar_lblMF(True)
    Else
        Call mostrar_lblMF(False)
    End If
    'ORCR FIN 20140414 *******
    'Jame Según TI-ERS037-2014 *******************************
    If RDatos!nPrdEstado = gColocEstRefNorm Or RDatos!nPrdEstado = gColocEstRefVenc Or RDatos!nPrdEstado = gColocEstRefMor Then
        cmdMotivoRefinanciado.Visible = True
    End If
    'Fin Jame ************************
    'ALPA20160623***********
    Dim oConstSistema As COMDConstSistema.DCOMConstSistema
    Set oConstSistema = New COMDConstSistema.DCOMConstSistema
    If oConstSistema.ObtenerVarSistemaCargoMultiple(gConstSistCuotaSistemaFinanciero, gsCodCargo) Then
        cmdDeudaCuoIFIS_CD.Visible = False
    Else
        cmdDeudaCuoIFIS_CD.Visible = True
    End If
    Set oConstSistema = Nothing
    '***********************
    Screen.MousePointer = 0
    Exit Function

ErrorCargaDatos:
        MsgBox err.Description, vbCritical, "Aviso"

End Function


Private Sub CmdAceptar_Click()
    Unload Me
End Sub

Private Sub cmdDeudaCuoIFIS_CD_Click()
Dim cPersCod As String
cPersCod = listaClientes.SelectedItem.SubItems(2)
Call frmCredEndeuCuotaSistFinanc.Inicio(ActxCta.NroCuenta, cPersCod)
End Sub

Private Sub cmdHistorial_Click()
    Dim oNCredDoc As COMNCredito.NCOMCredDoc
    Dim oPrevio As clsprevio
    Dim sCadImp As String
    
    Set oNCredDoc = New COMNCredito.NCOMCredDoc
    sCadImp = oNCredDoc.ImpreRepor_HistorialCredito(ActxCta.NroCuenta, gsNomAge, gdFecSis, gsCodUser, gsNomCmac)
    Set oNCredDoc = Nothing
    
    Set oPrevio = New clsprevio
    oPrevio.Show sCadImp, "Reporte de Historial del Credito", True
    Set oPrevio = Nothing
End Sub

Private Sub cmdIdentificacionRCC_Click()
Dim oCred As COMDCredito.DCOMCredDoc
Dim rs As ADODB.Recordset
Dim sCtaCod As String

sCtaCod = ActxCta.NroCuenta

'Solo para Mes y Consumo (ARCV 21-07-2006)
'If Mid(sCtaCod, 6, 1) <> Mid(gColPYMEEmp, 1, 1) And _
'    Mid(sCtaCod, 6, 1) <> Mid(gColConsuPlazoFijo, 1, 1) Then
'    MsgBox "Esta opción solo esta disponible para Creditos MES y CONSUMO", vbInformation, "Mensaje"
'    Exit Sub
'End If
If CInt(Mid(sCtaCod, 9, 1)) <> gMonedaExtranjera Then
    MsgBox "Esta opción solo esta disponible para Creditos en Dolares", vbInformation, "Mensaje"
    Exit Sub
End If

Set oCred = New COMDCredito.DCOMCredDoc

Set rs = oCred.ObtieneDatosIdentificacionRCC(sCtaCod)
Set oCred = Nothing

If Not rs.EOF Then
    Call ImprimeIdentificacionRCC(rs)
End If
End Sub

Sub ImprimeIdentificacionRCC(ByVal pRs As ADODB.Recordset)
    
    Dim fs As Scripting.FileSystemObject
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim nLineaInicio As Integer
    Dim nLineas As Integer
    Dim nLineasTemp As Integer
    
    Dim i As Integer
    Dim nTotal As Double
    
    Dim glsArchivo As String
    
    
    glsArchivo = "IdentificacionRCC_" & pRs!cPersCod & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    Set xlHoja1 = xlLibro.Worksheets.Add


    xlAplicacion.Range("A1:A1").ColumnWidth = 15
    xlAplicacion.Range("B1:B1").ColumnWidth = 10
    xlAplicacion.Range("C1:C1").ColumnWidth = 10
    xlAplicacion.Range("D1:D1").ColumnWidth = 12
    xlAplicacion.Range("E1:E1").ColumnWidth = 12
    xlAplicacion.Range("F1:F1").ColumnWidth = 12
                
    nLineas = 1
    xlHoja1.Cells(nLineas, 1) = "FORMATO PARA LA IDENTIFICACION DEL RIESGO CAMBIARIO CREDITICIO"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 6)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 1) = "CREDITOS MES O CONSUMO"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 6)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    nLineas = nLineas + 2
    xlHoja1.Cells(nLineas, 1) = "NOMBRE DEL CLIENTE"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    xlHoja1.Cells(nLineas, 2) = pRs!Titular
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 4)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders(xlEdgeBottom).LineStyle = 1

    xlHoja1.Cells(nLineas, 5) = pRs!documento
    xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineas, 5)).HorizontalAlignment = xlRight
    xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineas, 5)).Font.Bold = True
    xlHoja1.Cells(nLineas, 6) = pRs!NroDoc
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 1) = "AGENCIA"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    xlHoja1.Cells(nLineas, 2) = pRs!cAgeDescripcion
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders(xlEdgeBottom).LineStyle = 1
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 1) = "ACTIVIDAD"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    xlHoja1.Cells(nLineas, 2) = pRs!cCIIUdescripcion
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 4)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders(xlEdgeBottom).LineStyle = 1
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 1) = "MONTO DEL CREDITO"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    xlHoja1.Cells(nLineas, 2) = pRs!Moneda & CStr(Format(pRs!nMonto, "#,##0.00"))
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders(xlEdgeBottom).LineStyle = 1
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 1) = "PLAZO"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    xlHoja1.Cells(nLineas, 2) = lblPlazoAprobado & " MESES" 'CStr(prs!nPlazo)
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders(xlEdgeBottom).LineStyle = 1
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 1) = "ANALISTA DE CREDITOS"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    xlHoja1.Cells(nLineas, 2) = pRs!Analista
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders(xlEdgeBottom).LineStyle = 1
    
    nLineas = nLineas + 2
    xlHoja1.Cells(nLineas, 1) = "ESTRUCTURA DE LOS INGRESOS"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 6)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).HorizontalAlignment = xlCenter
    
    nLineas = nLineas + 2
    nLineasTemp = nLineas
    xlHoja1.Cells(nLineas, 2) = "MONEDA NACIONAL"
    xlHoja1.Cells(nLineas, 3) = "PORCENTAJE"
    xlHoja1.Cells(nLineas, 4) = "MONEDA EXTRANJERA"
    'xlHoja1.Cells(nLineas, 5) = "PORCENTAJE"
    'xlHoja1.Cells(nLineas, 6) = "TOTAL"
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 1) = "INGRESOS"
    xlHoja1.Cells(nLineas, 2) = pRs!IMN
    xlHoja1.Cells(nLineas, 3) = "100%"
    xlHoja1.Cells(nLineas, 4) = "0.00"
    'xlHoja1.Cells(nLineas, 5) = "0%"
    'xlHoja1.Cells(nLineas, 6) = prs!IMN
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 1) = "EGRESOS"
    xlHoja1.Cells(nLineas, 2) = pRs!EMN
    xlHoja1.Cells(nLineas, 3) = "100%"
    xlHoja1.Cells(nLineas, 4) = "0.00"
    'xlHoja1.Cells(nLineas, 5) = "0%"
    'xlHoja1.Cells(nLineas, 6) = prs!EMN
    xlHoja1.Range(xlHoja1.Cells(nLineasTemp, 1), xlHoja1.Cells(nLineas, 4)).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 2), xlHoja1.Cells(nLineas, 2)).NumberFormat = "#,##0.00"
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 4), xlHoja1.Cells(nLineas, 4)).NumberFormat = "#,##0.00"
    
    nLineas = nLineas + 2
    nLineasTemp = nLineas
    xlHoja1.Cells(nLineas, 1) = "SALDO DISPONIBLE EN S/."
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 2)).Merge True
    xlHoja1.Cells(nLineas, 3) = pRs!Saldo
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 1) = "MONTO DE LA CUOTA EN US $"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 2)).Merge True
    xlHoja1.Cells(nLineas, 3) = pRs!MontoCuotaDol
    'nLineas = nLineas + 1
    'xlHoja1.Cells(nLineas, 1) = "PATRIMONIO EXPRESADO EN S/."
    'xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 2)).Merge True
    'xlHoja1.Cells(nLineas, 3) = prs!Patrimonio
    'nLineas = nLineas + 1
    'xlHoja1.Cells(nLineas, 1) = "DEUDA SISTEMA FINANCIERO EN US $"
    'xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 2)).Merge True
    'xlHoja1.Cells(nLineas, 3) = prs!DeudaSF
    xlHoja1.Range(xlHoja1.Cells(nLineasTemp, 1), xlHoja1.Cells(nLineas, 3)).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineasTemp, 2), xlHoja1.Cells(nLineas, 4)).NumberFormat = "#,##0.00"
    
    'nLineas = nLineas + 2
    'nLineasTemp = nLineas
    'xlHoja1.Cells(nLineas, 1) = "POSICION DE CAMBIO"
    'xlHoja1.Cells(nLineas, 2) = "ACT ME - PAS ME"
    'nLineas = nLineas + 1
    'xlHoja1.Cells(nLineas, 1) = prs!PosicionCambio
    'xlHoja1.Cells(nLineas, 2) = prs!PosicionCambioValor
    'xlHoja1.Range(xlHoja1.Cells(nLineasTemp, 1), xlHoja1.Cells(nLineas, 2)).Borders.LineStyle = 1
    'xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).NumberFormat = "#,##0.00"
    
    'nLineasTemp = nLineas - 2
    'xlHoja1.Cells(nLineas - 2, 4) = "Inmueble $"
    'xlHoja1.Cells(nLineas - 1, 4) = "Maquinaria y Equipo"
    'xlHoja1.Cells(nLineas, 4) = "Otro Activo (mercaderia)"
    'xlHoja1.Range(xlHoja1.Cells(nLineas, 4), xlHoja1.Cells(nLineas, 5)).Merge True
    'nLineas = nLineas + 1
    'xlHoja1.Cells(nLineas, 4) = "Total Activo"
    'xlHoja1.Cells(nLineas, 5) = prs!TotalActivo
    xlHoja1.Range(xlHoja1.Cells(nLineasTemp, 4), xlHoja1.Cells(nLineas, 5)).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineas, 5), xlHoja1.Cells(nLineas, 5)).NumberFormat = "#,##0.00"
    
    nLineas = nLineas + 2
    xlHoja1.Cells(nLineas, 1) = "SIMULACION DE RIESGO CAMBIARIO"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 6)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).HorizontalAlignment = xlCenter
    
    nLineas = nLineas + 2
    nLineasTemp = nLineas
    xlHoja1.Cells(nLineas, 1) = "INDICADOR"
    xlHoja1.Cells(nLineas, 2) = "TIPO CAMBIO VIGENTE"
    xlHoja1.Cells(nLineas, 3) = "SHOCK 10%"
    xlHoja1.Cells(nLineas, 4) = "SHOCK 20%"
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas + 1, 1) = "VALOR CUOTA EN S/."
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas + 1, 1)).Merge True
    xlHoja1.Cells(nLineas, 2) = pRs!TipoCambio
    xlHoja1.Cells(nLineas, 3) = pRs!TipoCambio * 1.1
    xlHoja1.Cells(nLineas, 4) = pRs!TipoCambio * 1.2
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 2) = pRs!MontoCuotaSol
    xlHoja1.Cells(nLineas, 3) = pRs!MontoCuotaSol * 1.1
    xlHoja1.Cells(nLineas, 4) = pRs!MontoCuotaSol * 1.2
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 2), xlHoja1.Cells(nLineas, 4)).NumberFormat = "#,##0.00"
    
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 1) = "CUOTA/SALDO DISPONIBLE"
    Dim nPorcentaje As Double
    xlHoja1.Cells(nLineas, 2) = Format(CDbl(pRs!Cuota_Saldo) * 100, "0.00") '& " %"
    xlHoja1.Cells(nLineas, 3) = Format(pRs!Cuota_Saldo * 100 * 1.1, "0.00") '& "%"
    xlHoja1.Cells(nLineas, 4) = Format(pRs!Cuota_Saldo * 100 * 1.2, "0.00") '& "%"
    'nLineas = nLineas + 1
    'xlHoja1.Cells(nLineas, 1) = "VALOR DE LA DEUDA EN S/."
    'xlHoja1.Cells(nLineas, 2) = prs!DeudaSoles
    'xlHoja1.Cells(nLineas, 3) = prs!DeudaSoles * 1.1
    'xlHoja1.Cells(nLineas, 4) = prs!DeudaSoles * 1.2
    'nLineas = nLineas + 1
    'xlHoja1.Cells(nLineas, 1) = "DEUDA / PATRIMONIO"
    'xlHoja1.Cells(nLineas, 2) = prs!Deuda_Patrimonio
    'xlHoja1.Cells(nLineas, 3) = prs!Deuda_Patrimonio * 1.1
    'xlHoja1.Cells(nLineas, 4) = prs!Deuda_Patrimonio * 1.2
    xlHoja1.Range(xlHoja1.Cells(nLineasTemp, 1), xlHoja1.Cells(nLineas, 4)).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(nLineas - 1, 2), xlHoja1.Cells(nLineas, 4)).NumberFormat = "#,##0.00"
    
    nLineas = nLineas + 2
    nLineasTemp = nLineas
    xlHoja1.Cells(nLineas, 1) = "EXPUESTO"
    xlHoja1.Cells(nLineas, 2) = IIf(pRs!Expuesto = 1, "x", "")
    nLineas = nLineas + 1
    xlHoja1.Cells(nLineas, 1) = "NO EXPUESTO"
    xlHoja1.Cells(nLineas, 2) = IIf(pRs!Expuesto = 0, "x", "")
    xlHoja1.Range(xlHoja1.Cells(nLineasTemp, 1), xlHoja1.Cells(nLineas, 2)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(nLineasTemp, 1), xlHoja1.Cells(nLineas, 2)).Borders.LineStyle = 1
    
    nLineas = nLineas + 3
    xlHoja1.Cells(nLineas, 1) = "FIRMA ANALISTA"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    xlHoja1.Cells(nLineas, 2) = ""
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders(xlEdgeBottom).LineStyle = 1
    nLineas = nLineas + 2
    xlHoja1.Cells(nLineas, 1) = "FECHA DE EVALUACION"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    xlHoja1.Cells(nLineas, 2) = "CUSCO " & Day(gdFecSis) & " DE " & Format(gdFecSis, "MMMM") & " " & Year(gdFecSis)
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 3)).Borders(xlEdgeBottom).LineStyle = 1
    nLineas = nLineas + 2
    xlHoja1.Cells(nLineas, 1) = "VºBº ADMINISTRADOR"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    xlHoja1.Cells(nLineas, 2) = ""
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).Borders(xlEdgeBottom).LineStyle = 1
    nLineas = nLineas + 2
    xlHoja1.Cells(nLineas, 1) = "VºBº UNIDAD DE RIESGOS"
    xlHoja1.Range(xlHoja1.Cells(nLineas, 1), xlHoja1.Cells(nLineas, 1)).Font.Bold = True
    xlHoja1.Cells(nLineas, 2) = ""
    xlHoja1.Range(xlHoja1.Cells(nLineas, 2), xlHoja1.Cells(nLineas, 2)).Borders(xlEdgeBottom).LineStyle = 1
    
    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(2, 1)).Font.Size = 9
    xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(nLineas, 6)).Font.Size = 8
    xlHoja1.Cells.EntireColumn.AutoFit
    xlHoja1.Cells.EntireRow.AutoFit
    
    xlHoja1.SaveAs App.Path & "\SPOOLER\" & glsArchivo
               
    MsgBox "Se ha generado el Archivo en " & App.Path & "\SPOOLER\" & glsArchivo, vbInformation, "Mensaje"
    xlAplicacion.Visible = True
    xlAplicacion.Windows(1).Visible = True
        
    Set xlAplicacion = Nothing
    
End Sub

Private Sub cmdImprimir_Click()
Dim oCredDoc As COMNCredito.NCOMCredDoc
Dim Prev As previo.clsprevio
Dim MatRelacCred() As String
Dim MatDesembolsos() As String
Dim MatCuotasPend() As String
Dim MatHistorial(3, 7) As String
Dim MatDeudaVenc(6) As String
Dim MatDeudaAFecha(6) As String
Dim MatGarantias() As String
Dim MatRefinan() As Variant
Dim MatPagos() As String

Dim i As Integer
    On Error GoTo ErrorCmdImprimir_Click
    
    If Len(Trim(ActxCta.NroCuenta)) <> 18 Then
        MsgBox "No Existen Datos a Imprimir"
        Exit Sub
    End If
    
    MsgBox "Este documento no representa un Cronograma de Pagos oficial; es un documento interno.", vbInformation, "Aviso" 'JGPA20181011 ACTA 034-2018
    
    Screen.MousePointer = 11
    'Pago a la fecha
    MatDeudaAFecha(0) = lblSaldoKFecha.Caption
    MatDeudaAFecha(1) = CDbl(lblIntCompFecha.Caption) + CDbl(lblICV.Caption)
    MatDeudaAFecha(2) = lblIntMorFecha.Caption
    MatDeudaAFecha(3) = lblGastoFecha.Caption
    MatDeudaAFecha(4) = lblTotalFecha.Caption
    
    'Cuotas Vencidas
    MatDeudaVenc(0) = lblCapitalCuoPend.Caption
    MatDeudaVenc(1) = CDbl(lblInteresCuoPend.Caption) + CDbl(lblIntCompVenc.Caption)
    MatDeudaVenc(2) = lblMoraCuoPend.Caption
    MatDeudaVenc(3) = lblGastoCuoPend.Caption
    MatDeudaVenc(4) = lblTotalCuoPend.Caption
    
   'Carga Cuotas Pagadas
    'ReDim MatPagos(ListaPagos.ListItems.Count, 11)
        
    ' RIRO20131202 ERS098-2013 Historial Pago Credito
    Dim oPagos As New COMNCredito.NCOMCredDoc
    Dim rsPagos As ADODB.Recordset
    Set rsPagos = oPagos.HistorialPagoCreditos(ActxCta.NroCuenta)
    
    If Not rsPagos Is Nothing Then
        If Not rsPagos.EOF And Not rsPagos.BOF Then
            ReDim MatPagos(rsPagos.RecordCount, 12)
            For i = 0 To rsPagos.RecordCount - 1
                MatPagos(i, 0) = rsPagos!Fechapago 'ListaPagos.ListItems(i + 1).Text
                MatPagos(i, 1) = rsPagos!nNroCuota
                MatPagos(i, 2) = rsPagos!PagoTotal
                MatPagos(i, 3) = rsPagos!Capital
                MatPagos(i, 4) = rsPagos!Interes + rsPagos!ICV 'Se está incluyendo al ICV dentor del interés
                MatPagos(i, 5) = Format(rsPagos!Mora, "#0.00")
                MatPagos(i, 6) = rsPagos!Gastos
                MatPagos(i, 7) = Format(rsPagos!ITF, "#0.00")
                MatPagos(i, 8) = rsPagos!nDiasAtraso
                MatPagos(i, 9) = rsPagos!nSaldoCap
                MatPagos(i, 10) = Format(rsPagos!IntGra, "#0.00")
                rsPagos.MoveNext
            Next
        Else
            ReDim MatPagos(ListaPagos.ListItems.count, 12)
        End If
    Else
        ReDim MatPagos(ListaPagos.ListItems.count, 12)
    End If
    
    'Comentado Por RIRO20131202 Segun ERS098-2013
    ''Carga Cuotas Pagadas
    'ReDim MatPagos(ListaPagos.ListItems.Count, 10)
    'For i = 0 To ListaPagos.ListItems.Count - 1
    '    MatPagos(i, 0) = ListaPagos.ListItems(i + 1).Text
    '    MatPagos(i, 1) = ListaPagos.ListItems(i + 1).SubItems(1)
    '    MatPagos(i, 2) = ListaPagos.ListItems(i + 1).SubItems(2)
    '    MatPagos(i, 3) = ListaPagos.ListItems(i + 1).SubItems(3)
    '    MatPagos(i, 4) = ListaPagos.ListItems(i + 1).SubItems(4)
    '    MatPagos(i, 5) = ListaPagos.ListItems(i + 1).SubItems(5)
    '    MatPagos(i, 6) = ListaPagos.ListItems(i + 1).SubItems(6)
    '    MatPagos(i, 7) = ListaPagos.ListItems(i + 1).SubItems(7)
    '    MatPagos(i, 8) = ListaPagos.ListItems(i + 1).SubItems(8)
    '    MatPagos(i, 9) = ListaPagos.ListItems(i + 1).SubItems(9)
    'Next i
    
    'Carga Cuotas Pendientes
    ReDim MatCuotasPend(lstCuotasPend.ListItems.count, 9)
    For i = 0 To lstCuotasPend.ListItems.count - 1
        MatCuotasPend(i, 0) = lstCuotasPend.ListItems(i + 1).Text
        MatCuotasPend(i, 1) = lstCuotasPend.ListItems(i + 1).SubItems(1)
        MatCuotasPend(i, 2) = lstCuotasPend.ListItems(i + 1).SubItems(2)
        'MatCuotasPend(i, 3) = lstCuotasPend.ListItems(i + 1).SubItems(3)
        MatCuotasPend(i, 3) = (CDbl(lstCuotasPend.ListItems(i + 1).SubItems(3)) + _
                               CDbl(lstCuotasPend.ListItems(i + 1).SubItems(4)))
        MatCuotasPend(i, 4) = lstCuotasPend.ListItems(i + 1).SubItems(5)
        'MatCuotasPend(i, 5) = lstCuotasPend.ListItems(i + 1).SubItems(8)
        MatCuotasPend(i, 5) = lstCuotasPend.ListItems(i + 1).SubItems(13) 'EJVG20140925
        'MatCuotasPend(i, 6) = lstCuotasPend.ListItems(i + 1).SubItems(6)
        MatCuotasPend(i, 6) = lstCuotasPend.ListItems(i + 1).SubItems(11) 'EJVG20140925
        'MatCuotasPend(i, 7) = lstCuotasPend.ListItems(i + 1).SubItems(9)
        MatCuotasPend(i, 7) = lstCuotasPend.ListItems(i + 1).SubItems(14) 'EJVG20140925
        'MatCuotasPend(i, 8) = lstCuotasPend.ListItems(i + 1).SubItems(7)
        MatCuotasPend(i, 8) = lstCuotasPend.ListItems(i + 1).SubItems(12) 'EJVG20140925
        MatCuotasPend(i, 9) = lstCuotasPend.ListItems(i + 1).SubItems(5) 'RIRO20210503
        
    Next i
    
    'Carga Desembolsos
    ReDim MatDesembolsos(listaDesembolsos.ListItems.count, 4)
    For i = 0 To listaDesembolsos.ListItems.count - 1
        MatDesembolsos(i, 0) = listaDesembolsos.ListItems(i + 1).Text
        MatDesembolsos(i, 1) = listaDesembolsos.ListItems(i + 1).SubItems(2)
        MatDesembolsos(i, 2) = listaDesembolsos.ListItems(i + 1).SubItems(3)
        MatDesembolsos(i, 3) = listaDesembolsos.ListItems(i + 1).SubItems(4)
    Next i
    
    'Carga Relaciones de Clientes en Matriz
    ReDim MatRelacCred(listaClientes.ListItems.count, 2)
    For i = 0 To listaClientes.ListItems.count - 1
        MatRelacCred(i, 0) = PstaNombre(listaClientes.ListItems(i + 1).Text)
        MatRelacCred(i, 1) = PstaNombre(listaClientes.ListItems(i + 1).SubItems(1))
    Next i
    
    'Carga Garantias a Matriz
    ReDim MatGarantias(lstgarantias.ListItems.count, 6)
    For i = 0 To lstgarantias.ListItems.count - 1
        MatGarantias(i, 0) = lstgarantias.ListItems(i + 1).Text
        MatGarantias(i, 1) = lstgarantias.ListItems(i + 1).SubItems(1)
        MatGarantias(i, 2) = lstgarantias.ListItems(i + 1).SubItems(2)
        MatGarantias(i, 3) = lstgarantias.ListItems(i + 1).SubItems(3)
        MatGarantias(i, 4) = lstgarantias.ListItems(i + 1).SubItems(4)
        MatGarantias(i, 5) = lstgarantias.ListItems(i + 1).SubItems(5)
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
    
    ReDim MatRefinan(lstRefinanciados.ListItems.count, 4)
    For i = 1 To lstRefinanciados.ListItems.count
        MatRefinan(i - 1, 0) = lstRefinanciados.ListItems(i).Text
        MatRefinan(i - 1, 1) = lstRefinanciados.ListItems(i).SubItems(1)
        MatRefinan(i - 1, 2) = lstRefinanciados.ListItems(i).SubItems(2)
        MatRefinan(i - 1, 3) = lstRefinanciados.ListItems(i).SubItems(3)
    Next i
    
    Set oCredDoc = New COMNCredito.NCOMCredDoc
    Set Prev = New clsprevio

    ' *** RIRO20131202 Segun ERS098-2013 ***
    Dim nTasaEfectivaAnual, nTasaCostoEfectivoAnual As Double
    
    Dim oCred As COMNCredito.NCOMCredito
    Dim rTasCosEfeAnu As ADODB.Recordset
        
    Set oCred = New COMNCredito.NCOMCredito
            
    nTasaEfectivaAnual = oCred.TasaIntPerDias(Val(lbltasainteres.Caption), 360) * 100
    
    Dim oDCred As COMDCredito.DCOMCredActBD
    Set oDCred = New COMDCredito.DCOMCredActBD
    
    Set rTasCosEfeAnu = oDCred.CargaRecordSet("select nTasCosEfeAnu from ColocacCred where cCtaCod='" & ActxCta.NroCuenta & "'")
    
    If Not rTasCosEfeAnu Is Nothing Then
        If rTasCosEfeAnu.EOF And rTasCosEfeAnu.BOF Then
            nTasaCostoEfectivoAnual = 0
        Else
            nTasaCostoEfectivoAnual = IIf(IsNull(rTasCosEfeAnu!nTasCosEfeAnu), 0, rTasCosEfeAnu!nTasCosEfeAnu)
        End If
    Else
        nTasaCostoEfectivoAnual = 0
    End If
    
    Set oDCred = Nothing
    Set oCred = Nothing
    
    nTasaEfectivaAnual = Format(nTasaEfectivaAnual, "#0.00")
    nTasaCostoEfectivoAnual = Format(nTasaCostoEfectivoAnual, "#0.00")
    
    ' *** FIN RIRO ***
    Screen.MousePointer = 0

    'peac 20071228 se agrego "lblFecEEFF"
    'ORCR INICIO 20140414 ***
    Prev.Show oCredDoc.ImprimeConsultaCredito(ActxCta.NroCuenta, gsNomAge, gdFecSis, gsCodUser, lbltipoCredito.Caption, LblEstado _
        , lblLinea.Caption, lblfuente.Caption, MatRelacCred, lblanalista.Caption, lblcondicion.Caption, lbldestino.Caption, lblapoderado.Caption _
        , lbltipocuota.Caption, MatHistorial, MatDesembolsos, MatPagos, MatCuotasPend, MatDeudaVenc, MatDeudaAFecha, IIf(chkProtesto.value = 1, "SI", "NO"), IIf(chkCargoAuto.value = 1, "SI", "NO"), _
         IIf(chkRefinanciado.value = 1, "SI", "NO"), gsNomCmac, MatGarantias, ChkCuotaCom.value, ChkMiViv.value, ChkCalDin.value, LblIntCom.Caption, LblIntMor.Caption, MatRefinan, lblFecEEFF, nTasaEfectivaAnual, nTasaCostoEfectivoAnual, lbl_montoFinanciado.Caption, IIf(fbEsCredMIVIVIENDA = True, lblMensajeMIVIVIENDA.Caption, "")), "", True
    'WIOR 20151224 AGREGO IIf(fbEsCredMIVIVIENDA = True, lblMensajeMIVIVIENDA.Caption, "")
'    Prev.Show oCredDoc.ImprimeConsultaCredito(ActxCta.NroCuenta, gsNomAge, gdFecSis, gsCodUser, lbltipoCredito.Caption, lblEstado _
'        , lblLinea.Caption, lblfuente.Caption, MatRelacCred, lblanalista.Caption, lblcondicion.Caption, lbldestino.Caption, lblapoderado.Caption _
'        , lbltipocuota.Caption, MatHistorial, MatDesembolsos, MatPagos, MatCuotasPend, MatDeudaVenc, MatDeudaAFecha, IIf(chkProtesto.value = 1, "SI", "NO"), IIf(chkCargoAuto.value = 1, "SI", "NO"), _
'         IIf(chkRefinanciado.value = 1, "SI", "NO"), gsNomCmac, MatGarantias, ChkCuotaCom.value, ChkMiViv.value, ChkCalDin.value, LblIntCom.Caption, LblIntMor.Caption, MatRefinan, lblFecEEFF, 0, 0, lbl_montoFinanciado.Caption), "", True
    'ORCR FIN 20140414 ******
    
    Set Prev = Nothing
    Set oCredDoc = Nothing
    Exit Sub
ErrorCmdImprimir_Click:
    Screen.MousePointer = 0
    MsgBox err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub

Private Sub cmdMotivoRefinanciado_Click()
    Dim ofrmMotivoRefinanciamiento As New frmMotivoRefinanciamiento
    Call ofrmMotivoRefinanciamiento.Inicio(ActxCta.NroCuenta)
    Set ofrmMotivoRefinanciamiento = Nothing
End Sub

Private Sub CmdMuestra_Click()
'    Dim objDCRedito As COMDCredito.DCOMCredito
'    Dim rs As ADODB.Recordset
    Dim sCadImp As String
    Dim objNCredDoc As COMNCredito.NCOMCredDoc
'    Dim objDVisualizacion As COMNCredito.NCOMVisualizacion
    Dim objPrevio As clsprevio
'    Dim sTitular As String
'
'    Set objDCRedito = New COMDCredito.DCOMCredito
'    Set rs = objDCRedito.ListaKardex(ActxCta.NroCuenta)
'    Set objDCRedito = Nothing
'
'    Set objDVisualizacion = New COMNCredito.NCOMVisualizacion
'    sTitular = objDVisualizacion.ObtenerTitularByCredito(ActxCta.NroCuenta)
'    Set objDVisualizacion = Nothing
'
    Set objNCredDoc = New COMNCredito.NCOMCredDoc
    sCadImp = objNCredDoc.ImprimeKardexConsultaCredito(ActxCta.NroCuenta, gsNomAge, gdFecSis, gsCodUser, gsNomCmac)
    Set objNCredDoc = Nothing
    
    Set objPrevio = New clsprevio
    objPrevio.Show sCadImp, "Kardex del Credito"
    Set objPrevio = Nothing
    
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
    lblMonedaH.Caption = ""
    lblMonedaP.Caption = ""
    cmdMotivoRefinanciado.Visible = False 'JAME20140509
    lbl_Reprogramado.Visible = False 'FRHU20141031 OBSERVACION
    lblLiquidacion.Visible = False 'RIRO 20200911
    lblTotalFecha.ForeColor = vbBlack 'RIRO 20200911
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColPYMEEmp, False)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    CentraSdi Me
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    
'Identificacion RCC
Dim oParam As COMDCredito.DCOMParametro
Dim nValorParamRCC As Double
Set oParam = New COMDCredito.DCOMParametro
nValorParamRCC = oParam.RecuperaValorParametro(9004) '102733)
Set oParam = Nothing

If nValorParamRCC = 1 Then cmdIdentificacionRCC.Visible = True

lblMonedaP.Caption = ""
lblMonedaH.Caption = ""
gsOpeCod = gCredConsultaHistorialCred

'ORCR20140414 INICIO ***
lbl_montoFinanciado.Visible = False
lbl_MontoFinanciado_lbl.Visible = False
'ORCR20140414 FIN *******
End Sub

Private Sub listaClientes_DblClick()
    If listaClientes.ListItems.count > 0 Then
        If listaClientes.ListItems.count > 0 Then
            '*****RECO 20130701*******
            bVB = True
            Dim cPersCod As String
            cPersCod = listaClientes.SelectedItem.SubItems(2)
            frmBuscaPersona.VerificarEstadoPersona (cPersCod)
            
            If bVB = True Then
                Call frmPersona.Inicio(listaClientes.SelectedItem.SubItems(2), PersonaActualiza)
            End If
            '***END*******************
        End If
        '*****RECO 20130701*******
        'Call frmPersona.Inicio(listaClientes.SelectedItem.SubItems(2), PersonaActualiza)
        '***END*******************
    End If
End Sub

'Function ObtenerMoraVencida(ByVal pdHoy As Date, ByVal pMatCalend As Variant) As Double
'Dim i As Integer
'    ObtenerMoraVencida = 0
'    For i = 0 To UBound(pMatCalend) - 1
'        If pdHoy >= CDate(pMatCalend(i, 0)) Then
'            ObtenerMoraVencida = ObtenerMoraVencida + CDbl(pMatCalend(i, 6))
'        End If
'    Next i
'End Function
'Function ObtenerGastoVencido(ByVal pdHoy As Date, ByVal pMatCalend As Variant) As Double
'Dim i As Integer
'    ObtenerGastoVencido = 0
'    For i = 0 To UBound(pMatCalend) - 1
'        If pdHoy >= CDate(pMatCalend(i, 0)) Then
'            ObtenerGastoVencido = ObtenerGastoVencido + CDbl(pMatCalend(i, 9))
'        End If
'    Next i
'End Function

'Sub RecupCreditoAntiguo(ByVal psCtaCod As String)
'    Dim oDCredDoc As New COMDCredito.DCOMCredDoc
'    Dim sCtaCodAnt As String
'
'    Set oDCredDoc = New COMDCredito.DCOMCredDoc
'    sCtaCodAnt = oDCredDoc.Recup_CreditoAntiguo(psCtaCod)
'    Set oDCredDoc = Nothing
'
'    If Len(sCtaCodAnt) = 0 Then
'        LblAntigua = "ES UN CREDITO NUEVO"
'    Else
'        LblAntigua.Caption = sCtaCodAnt
'    End If
'
'
'End Sub


'Sub Recup_FechaCancelacion(ByVal psCtaCod As String)
'    Dim oDCredDoc As COMDCredito.DCOMCredDoc
'    Dim sFecha As String
'
'    Set oDCredDoc = New COMDCredito.DCOMCredDoc
'    sFecha = oDCredDoc.Recup_FechaCancelacion(psCtaCod)
'    Set oDCredDoc = Nothing
'    txtFechaCancelacion.Text = sFecha
'End Sub

'JACA 20120112************************************************************************
Private Sub lstCuotasPend_Click()
    
    'Al hacer click en una cuota pendiente,llenara datos en la pestaña Comp. Pago
    If lstCuotasPend.ListItems.count > 0 Then
        LimpiarFecCompromisoPago
        Me.lblCompCuota = lstCuotasPend.ListItems.Item(lstCuotasPend.SelectedItem.Index)
        'Me.lblCompCuotaSaldo = lstCuotasPend.ListItems.iTem(lstCuotasPend.SelectedItem.Index).SubItems(7)
        Me.lblCompCuotaSaldo = lstCuotasPend.ListItems.Item(lstCuotasPend.SelectedItem.Index).SubItems(11) 'EJVG20140925
        Me.lblCompCuotaFecha = lstCuotasPend.ListItems.Item(lstCuotasPend.SelectedItem.Index).SubItems(1)
    End If

End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    
    'Al hacer click en la pestaña Comp.Pago
    If Me.SSTab1.Tab = 7 Then
        If Me.lblCompCuota = "" Then
            Me.lblCompMsg = "Seleccione una Cuota de la Pestaña de Cuotas_Pendientes"
        Else
            CalcularDiasCompromisoPago True
        End If
    Else
        Me.lblCompMsg = ""
    End If
End Sub

Private Sub txtCompFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.lblCompCuotaFecha <> "" Then
        CalcularDiasCompromisoPago
        If cmdCompGuardar.Enabled = True Then
            Me.cmdCompGuardar.SetFocus
        End If
    End If
End Sub
Private Sub CalcularDiasCompromisoPago(Optional bgetFecha As Boolean = False)
    
    Dim rs As Recordset
    Dim oNCredito As COMNCredito.NCOMCredito
    Dim sFechaMsg As String
    
    'Si hace en la pestaña de Com_Pago; Obtiene Fecha de Pago Guardado
    If bgetFecha Then
        Set rs = New Recordset
        Set oNCredito = New COMNCredito.NCOMCredito
        Set rs = oNCredito.obtenerFechaCompromisoPago(Me.ActxCta.NroCuenta, Me.lblCompCuota)
        
        If Not (rs.EOF And rs.BOF) Then
            Me.txtCompFecha.Text = rs!dFecCompPago
        Else
            Exit Sub
        End If
        
        Set rs = Nothing
        Set oNCredito = Nothing
   
    Else  'Solo cuando hace Enter en la celda de fecha
            sFechaMsg = ValidaFecha(Me.txtCompFecha.Text)
            If sFechaMsg <> "" Then
                MsgBox sFechaMsg, vbInformation, "AVISO!"
                Exit Sub
            End If
    End If
    
    'Calcula los dias faltantes o pasados de la fecha actual con la fecha de compromiso
    If DateDiff("d", Me.txtCompFecha.Text, gdFecSis) <= 0 Then
            Me.lblCompDiasMsg = "Nro Dias Faltantes:"
            Me.lblCompDias.ForeColor = &H80000008
    Else
            Me.lblCompDiasMsg = "Nro Dias Pasados:"
            Me.lblCompDias.ForeColor = &HFF&
    End If
    Me.lblCompDias = Abs(DateDiff("d", Me.txtCompFecha.Text, gdFecSis))
       
    
End Sub
Private Sub cmdCompGuardar_Click()
    Dim oNCredito As COMNCredito.NCOMCredito
    Dim sFechaMsg As String
    Call txtCompFecha_KeyPress(13) 'ALPA 20120305
    sFechaMsg = ValidaFecha(Me.txtCompFecha.Text)
    
    If sFechaMsg <> "" Then
        MsgBox sFechaMsg, vbInformation, "AVISO!"
        Exit Sub
    End If
    
    Set oNCredito = New COMNCredito.NCOMCredito
    
    If MsgBox("Se va ha Guardar la Fecha de Compromiso de Pago", vbYesNo, "Guardar Fecha de Compromiso") = vbYes Then
    
        Dim sMovNro As String
        Dim clsMov As New COMNContabilidad.NCOMContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
        If oNCredito.guardarFechaCompromisoPago(Me.ActxCta.NroCuenta, Me.lblCompCuota, Me.txtCompFecha.Text, sMovNro) Then
           
            MsgBox "Se han Guardado los Datos con Exito!", vbInformation, "Datos Guardado"
        Else
            MsgBox "No se Pudieron Guardar los Datos!", vbInformation, "Error al Guardar"
        End If
        
        Set clsMov = Nothing
    End If
    
    Set oNCredito = Nothing
End Sub
Private Sub LimpiarFecCompromisoPago()
    lblCompCuota = ""
    lblCompCuotaSaldo = ""
    lblCompCuotaFecha = ""
    lblCompMsg = ""
    txtCompFecha.Text = "__/__/____"
    lblCompDias = ""
    Me.lblCompDiasMsg = "Nro Dias Faltantes:"
    Me.lblCompDias.ForeColor = &H80000008
End Sub
Private Sub MostrarFecCompromisoPago()
    Dim oNCredito As New COMNCredito.NCOMCredito
    Dim oDCredito As New COMDCredito.DCOMCredito
    
    'Verfica si tiene Permiso
    If oNCredito.verificarPermisoFechaCompPago(gsCodCargo, gsCodPersUser) Then
        SSTab1.TabVisible(7) = True
        
        If oDCredito.GetCargoActualizarCompPago(ActxCta.NroCuenta, gsCodPersUser) Then
            cmdCompGuardar.Enabled = True
        Else
            cmdCompGuardar.Enabled = False
        End If
        
    Else
         SSTab1.TabVisible(7) = False
    End If

End Sub
'JACA END************************************************************************
'ORCR INICIO 20140414 ***
Private Sub mostrar_lblMF(Mostrar As Boolean)
    lbl_montoFinanciado.Visible = Mostrar
    lbl_MontoFinanciado_lbl.Visible = Mostrar
End Sub
'ORCR FIN 20140414 ***
'EJVG20160308 ***
Private Sub lstgarantias_DblClick()
    If lstgarantias.ListItems.count > 0 Then
        frmGarantia.Consultar lstgarantias.SelectedItem.SubItems(6)
    End If
End Sub
'END EJVG *******
