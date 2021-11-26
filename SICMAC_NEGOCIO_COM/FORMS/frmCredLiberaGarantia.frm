VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCredLiberaGarantia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liberar Garantias"
   ClientHeight    =   6885
   ClientLeft      =   3330
   ClientTop       =   2040
   ClientWidth     =   9015
   Icon            =   "frmCredLiberaGarantia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdNuevaBuesq 
      Caption         =   "&Nueva Busqueda"
      Height          =   435
      Left            =   5985
      TabIndex        =   39
      Top             =   6345
      Width           =   1575
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   7575
      TabIndex        =   38
      Top             =   6345
      Width           =   1260
   End
   Begin TabDlg.SSTab SSDatos 
      Height          =   4530
      Left            =   60
      TabIndex        =   14
      Top             =   1740
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   7990
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Garantias"
      TabPicture(0)   =   "frmCredLiberaGarantia.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(2)=   "DGGarantias"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Datos"
      TabPicture(1)   =   "frmCredLiberaGarantia.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "LblGarBanco"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "framontos"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "FraClase"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "FraTipoRea"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame4"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame5"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Creditos"
      TabPicture(2)   =   "frmCredLiberaGarantia.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label14"
      Tab(2).Control(1)=   "Label12"
      Tab(2).Control(2)=   "DGCredito"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame5 
         Enabled         =   0   'False
         Height          =   600
         Left            =   5685
         TabIndex        =   44
         Top             =   3090
         Width           =   2790
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
            Left            =   465
            TabIndex        =   45
            Top             =   195
            Width           =   1620
         End
      End
      Begin VB.Frame Frame4 
         Height          =   630
         Left            =   300
         TabIndex        =   35
         Top             =   3705
         Width           =   7950
         Begin VB.CommandButton CmdDesbGarant 
            Caption         =   "&Desbloquear Garantia"
            Enabled         =   0   'False
            Height          =   360
            Left            =   3315
            TabIndex        =   53
            Top             =   180
            Width           =   1995
         End
         Begin VB.CommandButton CmdBloquear 
            Caption         =   "&Bloquear Garantia"
            Enabled         =   0   'False
            Height          =   360
            Left            =   1695
            TabIndex        =   50
            Top             =   180
            Width           =   1590
         End
         Begin VB.CommandButton CmdLiberaGarantia 
            Caption         =   "&Liberar Garantia"
            Enabled         =   0   'False
            Height          =   360
            Left            =   60
            TabIndex        =   36
            Top             =   180
            Width           =   1605
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1305
         Left            =   150
         TabIndex        =   28
         Top             =   435
         Width           =   8655
         Begin VB.Label LblMoneda 
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
            Height          =   300
            Left            =   1620
            TabIndex        =   58
            Top             =   915
            Width           =   1380
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Moneda :"
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
            Left            =   195
            TabIndex        =   57
            Top             =   975
            Width           =   810
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Estado :"
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
            Left            =   5985
            TabIndex        =   52
            Top             =   300
            Width           =   720
         End
         Begin VB.Label LblEstado 
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
            Height          =   300
            Left            =   6870
            TabIndex        =   51
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label LblgarNumDoc 
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
            Height          =   300
            Left            =   6870
            TabIndex        =   34
            Top             =   585
            Width           =   1380
         End
         Begin VB.Label LblGarDoc 
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
            Height          =   300
            Left            =   1620
            TabIndex        =   33
            Top             =   585
            Width           =   4260
         End
         Begin VB.Label LblGarTpo 
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
            Height          =   300
            Left            =   1620
            TabIndex        =   32
            Top             =   255
            Width           =   4260
         End
         Begin VB.Label Label3 
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
            Left            =   180
            TabIndex        =   31
            Top             =   285
            Width           =   1275
         End
         Begin VB.Label Label2 
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
            Left            =   195
            TabIndex        =   30
            Top             =   645
            Width           =   1230
         End
         Begin VB.Label Label1 
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
            Left            =   5985
            TabIndex        =   29
            Top             =   645
            Width           =   810
         End
      End
      Begin VB.Frame FraTipoRea 
         Caption         =   "Tipo de Realizacion"
         Enabled         =   0   'False
         Height          =   615
         Left            =   270
         TabIndex        =   24
         Top             =   2505
         Width           =   4020
         Begin VB.OptionButton OptTR 
            Caption         =   "De Lenta Realizacion"
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   26
            Top             =   255
            Value           =   -1  'True
            Width           =   1950
         End
         Begin VB.OptionButton OptTR 
            Caption         =   "De Rapida Realizacion"
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   2010
            TabIndex        =   25
            Top             =   270
            Width           =   1980
         End
      End
      Begin VB.Frame FraClase 
         Caption         =   "Clase de Garantia"
         Enabled         =   0   'False
         Height          =   615
         Left            =   270
         TabIndex        =   21
         Top             =   1815
         Width           =   4005
         Begin VB.OptionButton OptCG 
            Caption         =   "Garantia Preferida"
            Height          =   240
            Index           =   1
            Left            =   2025
            TabIndex        =   23
            Top             =   255
            Width           =   1650
         End
         Begin VB.OptionButton OptCG 
            Caption         =   "Garantia No Preferida"
            Height          =   240
            Index           =   0
            Left            =   105
            TabIndex        =   22
            Top             =   255
            Value           =   -1  'True
            Width           =   1905
         End
      End
      Begin VB.Frame framontos 
         Height          =   1335
         Left            =   5685
         TabIndex        =   16
         Top             =   1700
         Width           =   2775
         Begin VB.Label LblMontoxGrav 
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
            ForeColor       =   &H8000000D&
            Height          =   300
            Left            =   1350
            TabIndex        =   43
            Top             =   915
            Width           =   1260
         End
         Begin VB.Label LblMontoRea 
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
            ForeColor       =   &H8000000D&
            Height          =   300
            Left            =   1350
            TabIndex        =   42
            Top             =   600
            Width           =   1260
         End
         Begin VB.Label LblMontoTas 
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
            ForeColor       =   &H8000000D&
            Height          =   300
            Left            =   1350
            TabIndex        =   41
            Top             =   285
            Width           =   1260
         End
         Begin VB.Label lbltasa 
            AutoSize        =   -1  'True
            Caption         =   "Valor Comercial :"
            Height          =   195
            Left            =   75
            TabIndex        =   20
            ToolTipText     =   "Monto Tasación"
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label lblrealizacion 
            AutoSize        =   -1  'True
            Caption         =   "Realización :"
            Height          =   195
            Left            =   360
            TabIndex        =   19
            Top             =   630
            Width           =   915
         End
         Begin VB.Label lblMontoGrav 
            AutoSize        =   -1  'True
            Caption         =   "Disponible :"
            Height          =   195
            Left            =   440
            TabIndex        =   18
            Top             =   960
            Width           =   825
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
            TabIndex        =   17
            ToolTipText     =   "Monto Tasación"
            Top             =   15
            Width           =   630
         End
      End
      Begin MSDataGridLib.DataGrid DGGarantias 
         Height          =   3405
         Left            =   -74835
         TabIndex        =   15
         Top             =   675
         Width           =   8550
         _ExtentX        =   15081
         _ExtentY        =   6006
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "cEstadoGar"
            Caption         =   "Estado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cTpoGarDescripcion"
            Caption         =   "Garantia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cNroDoc"
            Caption         =   "Documento"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cDescripcion"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "cPersNombre"
            Caption         =   "Emisor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3254.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3404.977
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   5130.142
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DGCredito 
         Height          =   3405
         Left            =   -74775
         TabIndex        =   56
         Top             =   675
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   6006
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "cCtaCod"
            Caption         =   "Credito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cMoneda"
            Caption         =   "Moneda"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "nGravado"
            Caption         =   "Gravamen"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "nSaldo"
            Caption         =   "Saldo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "nCuotas"
            Caption         =   "Cuotas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "nNroProxCuota"
            Caption         =   "ProxCuota"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   2115.213
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1049.953
            EndProperty
         EndProperty
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Creditos"
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
         Height          =   195
         Left            =   -67935
         TabIndex        =   54
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Garantias"
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
         Height          =   195
         Left            =   -73965
         TabIndex        =   49
         Top             =   105
         Width           =   1080
      End
      Begin VB.Label Label6 
         Caption         =   "Garantias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   -73950
         TabIndex        =   48
         Top             =   75
         Width           =   1080
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Datos"
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
         Height          =   195
         Left            =   4230
         TabIndex        =   47
         Top             =   105
         Width           =   525
      End
      Begin VB.Label LblGarBanco 
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
         Height          =   300
         Left            =   300
         TabIndex        =   37
         Top             =   3390
         Width           =   3975
      End
      Begin VB.Label Label7 
         Caption         =   "Banco"
         Height          =   255
         Left            =   315
         TabIndex        =   27
         Top             =   3150
         Width           =   555
      End
      Begin VB.Label Label4 
         Caption         =   "Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4245
         TabIndex        =   46
         Top             =   75
         Width           =   525
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Creditos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   -67920
         TabIndex        =   55
         Top             =   105
         Width           =   705
      End
   End
   Begin VB.Frame FraBuscar 
      Caption         =   "Busqueda Garantia"
      Height          =   1605
      Left            =   810
      TabIndex        =   0
      Top             =   60
      Width           =   7575
      Begin VB.Frame FraBusqGar 
         Height          =   1395
         Left            =   1755
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   4950
         Begin VB.CommandButton CmdGarBuscar 
            Caption         =   "&Buscar"
            Height          =   345
            Left            =   3315
            TabIndex        =   40
            Top             =   930
            Width           =   1380
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
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipos de Garantias"
            Top             =   195
            Width           =   3285
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
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Tag             =   "cboPrincipal"
            ToolTipText     =   "Tipo de Documentos"
            Top             =   570
            Width           =   3285
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
            Left            =   1410
            MaxLength       =   15
            TabIndex        =   8
            Tag             =   "txtPrincipal"
            Top             =   930
            Width           =   1500
         End
         Begin VB.Label Label9 
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
            Left            =   105
            TabIndex        =   13
            Top             =   255
            Width           =   1275
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
            Left            =   105
            TabIndex        =   12
            Top             =   630
            Width           =   1230
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
            Left            =   105
            TabIndex        =   11
            Top             =   960
            Width           =   810
         End
      End
      Begin VB.Frame Frame2 
         Height          =   825
         Left            =   90
         TabIndex        =   1
         Top             =   420
         Width           =   1605
         Begin VB.OptionButton OptBusq 
            Caption         =   "Por Garantia"
            Height          =   315
            Index           =   1
            Left            =   75
            TabIndex        =   3
            Top             =   420
            Width           =   1185
         End
         Begin VB.OptionButton OptBusq 
            Caption         =   "Por Titular"
            Height          =   315
            Index           =   0
            Left            =   75
            TabIndex        =   2
            Top             =   150
            Value           =   -1  'True
            Width           =   1140
         End
      End
      Begin VB.Frame FraBusqTitu 
         Height          =   840
         Left            =   1755
         TabIndex        =   4
         Top             =   450
         Width           =   5700
         Begin VB.CommandButton CmdBusqTitu 
            Caption         =   "&Buscar"
            Height          =   390
            Left            =   4425
            TabIndex        =   5
            Top             =   270
            Width           =   1050
         End
         Begin VB.Label LblBusqTitu 
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
            Height          =   300
            Left            =   105
            TabIndex        =   6
            Top             =   300
            Width           =   4260
         End
      End
   End
End
Attribute VB_Name = "frmCredLiberaGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R As ADODB.Recordset
Dim RCred As ADODB.Recordset
Dim objPista As COMManejador.Pista


Private Sub CargaControles()
Dim oGarantia As COMDCredito.DCOMGarantia
Dim rsTipoGar As ADODB.Recordset
Dim rsDocum As ADODB.Recordset

    Set oGarantia = New COMDCredito.DCOMGarantia
    Call oGarantia.CargarControlesLiberaGarantia(rsDocum, rsTipoGar)
    Set oGarantia = Nothing

    'Carga Tipo Garantias
    CmbTipoGarant.Clear
    Call CambiaTamañoCombo(CmbTipoGarant)
    'Call CargaComboConstante(gPersGarantia, CmbTipoGarant)
    Call Llenar_Combo_con_Recordset(rsTipoGar, CmbTipoGarant)
        
    CmbDocGarant.Clear
'        Set oGarantia = New COMDCredito.DCOMGarantia
'        Set R = oGarantia.RecuperaTiposDocumentosGarantia()
'        Set oGarantia = Nothing
        CmbDocGarant.Clear
        Do While Not rsDocum.EOF
            CmbDocGarant.AddItem rsDocum!cDocDesc & Space(150) & rsDocum!nDocTpo
            rsDocum.MoveNext
        Loop
        'rsDocum.Close
        
    Call CambiaTamañoCombo(CmbDocGarant, 300)
    
End Sub

Private Sub LimpiaPantalla()
    Call LimpiaControles(Me)
    Set DGGarantias.DataSource = Nothing
    DGGarantias.Refresh
    FraBuscar.Enabled = True
    If OptBusq(0).value = True Then
        CmdBusqTitu.SetFocus
    Else
        CmdGarBuscar.SetFocus
    End If
    CmdBloquear.Enabled = False
    CmdLiberaGarantia.Enabled = False
    CmdDesbGarant.Enabled = False
    SSDatos.Tab = 0
    Set DGCredito.DataSource = Nothing
End Sub

Private Sub CmbTipoGarant_Click()
Dim oGarantia As COMDCredito.DCOMGarantia
    
        CmbDocGarant.Clear
        If CmbTipoGarant.Text = "" Then
            Exit Sub
        End If
        Set oGarantia = New COMDCredito.DCOMGarantia
        Set R = oGarantia.RecuperaTiposDocumGarantias(CInt(Trim(Right(CmbTipoGarant.Text, 10))))
        Set oGarantia = Nothing
        CmbDocGarant.Clear
        Do While Not R.EOF
            CmbDocGarant.AddItem R!cDocDesc & Space(150) & R!nDocTpo
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
    Call CambiaTamañoCombo(CmbDocGarant, 300)
End Sub

Private Sub CmdBloquear_Click()
Dim oGar As COMNCredito.NCOMGarantia
Dim odGar As COMDCredito.DCOMGarantia
Dim bResul As Boolean
        

    If R!nEstado = gPersGarantEstadoBloqueada Then
        MsgBox "La Garantia ya esta Bloqueada", vbInformation, "Aviso"
        Exit Sub
    End If

    Set odGar = New COMDCredito.DCOMGarantia
    bResul = odGar.PerteneceAGarantiaPendienteAsiento(R!cNumGarant, gdFecSis)
    Set odGar = Nothing
    
    If bResul Then
        MsgBox "No se Puede Bloquear la Garantia, porque esta asociada a un Credito recien Desembolsado", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Al Bloquear la Garantia No Podra ser Usada por Otros Creditos, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oGar = New COMNCredito.NCOMGarantia
    Call oGar.BloqueaGarantia(R!cNumGarant)
    
    ''*** PEAC 20090126
    objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Bloquear garantia.", R!cNumGarant, gCodigoGarantia
    
    Set oGar = Nothing
    
    CmdNuevaBuesq_Click
    
End Sub

Private Sub CmdBusqTitu_Click()
Dim oPers As COMDPersona.UCOMPersona ' UPersona
Dim oGar As COMDCredito.DCOMGarantia

    Set oPers = frmBuscaPersona.Inicio
    If oPers Is Nothing Then
        MsgBox "No se encontraron Datos", vbInformation, "Aviso"
        Exit Sub
    End If
    If oPers.sPersCod <> "" Then
        Set oGar = New COMDCredito.DCOMGarantia
        Set R = oGar.RecuperaGarantiasPersona(oPers.sPersCod, , True)
        Set DGGarantias.DataSource = R
        Set oGar = Nothing
    End If
    
    If R.RecordCount = 0 Then
            MsgBox "No se encontraron Datos", vbInformation, "Aviso"
            CmdBusqTitu.SetFocus
    Else
            LblBusqTitu.Caption = oPers.sPersNombre
            DGGarantias.SetFocus
            FraBuscar.Enabled = False
    End If
    Set oPers = Nothing
End Sub

Private Sub CmdDesbGarant_Click()
Dim oGar As COMNCredito.NCOMGarantia

    If R!nEstado <> gPersGarantEstadoBloqueada Then
        MsgBox "La Garantia ya esta Desbloqueada", vbInformation, "Aviso"
        Exit Sub
    End If

    If MsgBox("Se va a Desbloquear la Garantia y podra ser Usada por Otros Creditos, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oGar = New COMNCredito.NCOMGarantia
    Call oGar.DesbloqueaGarantia(R!cNumGarant)
    
    ''*** PEAC 20090126
    objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Desbloquear garantia.", R!cNumGarant, gCodigoGarantia
    
    Set oGar = Nothing
    
    CmdNuevaBuesq_Click
    
End Sub

Private Sub CmdGarBuscar_Click()
Dim oGar As COMDCredito.DCOMGarantia

        If CmbTipoGarant.Text = "" Then
            MsgBox "Ingrese el Tipo de Garantia"
            Exit Sub
        End If
        Set oGar = New COMDCredito.DCOMGarantia
        Set R = oGar.RecuperaGarantiasxDatos(Trim(Right(CmbTipoGarant.Text, 10)), Trim(Right(CmbDocGarant.Text, 10)), Trim(txtNumDoc.Text))
        Set DGGarantias.DataSource = R
        Set oGar = Nothing
        If R.RecordCount = 0 Then
            MsgBox "No se encontraron Datos", vbInformation, "Aviso"
            CmbTipoGarant.SetFocus
        Else
            DGGarantias.SetFocus
            FraBuscar.Enabled = False
        End If
End Sub

Private Sub CmdLiberaGarantia_Click()
Dim oGar As COMNCredito.NCOMGarantia
Dim Pos As Integer
Dim sCodCta As String
Dim pnEstado As Integer
Dim odGar As COMDCredito.DCOMGarantia
Dim bResul As Boolean
        
    'ARCV 05-07-2007
    'Set odGar = New COMDCredito.DCOMGarantia
    'bResul = odGar.PerteneceAGarantiaPendienteAsiento(R!cNumGarant, gdFecSis)
    'Set odGar = Nothing
    '
    'If bResul Then
    '    MsgBox "No se Puede Liberar la Garantia, porque esta asociada a un Credito recien Desembolsado", vbInformation, "Aviso"
    '    Exit Sub
    'End If
    '-------
    If RCred.EOF Then
        MsgBox "No hay creditos asignados a la Garantia", vbInformation, "Mensaje"
        Exit Sub
    End If
    sCodCta = RCred!cCtaCod
    If RCred.RecordCount > 1 Then
        pnEstado = gPersGarantEstadoContabilizado
    Else
        pnEstado = gPersGarantEstadoLiberado
    End If
    'If R!nEstado <> gPersGarantEstadoContabilizado Then
    '    MsgBox "Solo puede Liberar Garantias que estan en estado Contabilizada", vbInformation, "Aviso"
    '    Exit Sub
    'End If
    sCodCta = RCred!cCtaCod
    Set oGar = New COMNCredito.NCOMGarantia
    Call oGar.LiberaGarantia(R!cNumGarant, R!nTpoGarantia, RCred!nGravado, R!nmoneda, gsCodAge, gsCodUser, gdFecSis, sCodCta, pnEstado)
    
    ''*** PEAC 20090126
    objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Liberar garantia.", R!cNumGarant, gCodigoGarantia
    
    Set oGar = Nothing
    
    Call CmdNuevaBuesq_Click
    
End Sub

Private Sub CmdNuevaBuesq_Click()
    Call LimpiaPantalla
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub


Private Sub DGGarantias_Click()
Dim oGar As COMDCredito.DCOMGarantia

    If R Is Nothing Then
        Exit Sub
    End If
    If R.RecordCount > 0 Then
        LblGarTpo.Caption = Trim(R!cTpoGarDescripcion)
        LblGarDoc.Caption = Trim(R!cDocDesc)
        LblgarNumDoc.Caption = Trim(R!cNroDoc)
        If R!nGarClase = 0 Then
            OptCG(0).value = True
        Else
            OptCG(0).value = True
        End If
        If R!nGarTpoRealiz = 0 Then
            OptTR(0).value = True
        Else
            OptTR(1).value = False
        End If
        
        LblGarBanco.Caption = PstaNombre(IIf(IsNull(R!cBanco), "", R!cBanco))
        LblMontoTas.Caption = Format(R!nTasacion, "#0.00")
        LblMontoRea.Caption = Format(R!nRealizacion, "#0.00")
        LblMontoxGrav.Caption = Format(R!nDisponible, "#0.00")
        ChkGarReal.value = R!nGarantReal
        lblEstado.Caption = Trim(R!cEstadoGar)
        lblMoneda.Caption = Trim(R!cMonedaDesc)
        CmdBloquear.Enabled = True
        CmdLiberaGarantia.Enabled = True
        CmdDesbGarant.Enabled = True
        
        Set oGar = New COMDCredito.DCOMGarantia
         Set RCred = oGar.RecuperaGarantiaCreditoDatosVigente(R!cNumGarant)
        Set DGCredito.DataSource = RCred
        Set oGar = Nothing
    End If
End Sub


Private Sub DGGarantias_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call DGGarantias_Click
End Sub

Private Sub Form_Load()
    CentraForm Me
    CargaControles
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredLiberarBloquearGarantia
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Private Sub OptBusq_Click(Index As Integer)
    If Index = 0 Then
        FraBusqTitu.Visible = True
        FraBusqGar.Visible = False
        LblBusqTitu.Caption = ""
    Else
        FraBusqTitu.Visible = False
        FraBusqGar.Visible = True
        CmbTipoGarant.ListIndex = -1
    End If
    LimpiaPantalla
End Sub

