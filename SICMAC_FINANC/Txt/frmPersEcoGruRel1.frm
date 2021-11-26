VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPersEcoGruRel1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MANTENIMIENTO DE GRUPOS ECONOMICOS"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   Icon            =   "frmPersEcoGruRel1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9120
      TabIndex        =   60
      Top             =   7410
      Width           =   1185
   End
   Begin VB.PictureBox PG2B 
      Height          =   1635
      Left            =   45
      ScaleHeight     =   1575
      ScaleWidth      =   5820
      TabIndex        =   34
      Top             =   6405
      Visible         =   0   'False
      Width           =   5880
      Begin VB.CheckBox chkEcoc 
         Caption         =   "Gestión"
         Height          =   240
         Left            =   4395
         TabIndex        =   56
         Top             =   780
         Width           =   1560
      End
      Begin VB.CheckBox chkEcob 
         Caption         =   "Prop. Indirecta"
         Height          =   240
         Left            =   4380
         TabIndex        =   55
         Top             =   465
         Width           =   1560
      End
      Begin VB.CheckBox chkEcoa 
         Caption         =   "Prop. Directa"
         Height          =   240
         Left            =   4380
         TabIndex        =   54
         Top             =   150
         Width           =   1560
      End
      Begin VB.TextBox txtTexto 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   885
         MaxLength       =   30
         TabIndex        =   53
         Top             =   780
         Width           =   3405
      End
      Begin VB.CommandButton cmdGrabarG2 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3330
         TabIndex        =   20
         Top             =   1200
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancelarG2 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4575
         TabIndex        =   21
         Top             =   1200
         Width           =   1185
      End
      Begin VB.ComboBox CBOG2 
         Height          =   315
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1155
         Width           =   2280
      End
      Begin VB.TextBox txtPersNombreG2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   18
         Top             =   427
         Width           =   4155
      End
      Begin Sicmact.TxtBuscar txtPersCodG2 
         Height          =   330
         Left            =   2400
         TabIndex        =   17
         Top             =   45
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCodigo3 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   180
         TabIndex        =   51
         Top             =   1365
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lblCodigo2 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   165
         TabIndex        =   47
         Top             =   1080
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Relación"
         Height          =   195
         Left            =   165
         TabIndex        =   36
         Top             =   810
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   150
         TabIndex        =   35
         Top             =   60
         Width           =   555
      End
   End
   Begin VB.PictureBox PG3B 
      Height          =   1635
      Left            =   60
      ScaleHeight     =   1575
      ScaleWidth      =   5820
      TabIndex        =   39
      Top             =   6420
      Visible         =   0   'False
      Width           =   5880
      Begin VB.CheckBox chkEco1 
         Caption         =   "Prop. Directa"
         Height          =   240
         Left            =   4380
         TabIndex        =   59
         Top             =   150
         Width           =   1410
      End
      Begin VB.CheckBox chkEco2 
         Caption         =   "Prop. Indirecta"
         Height          =   240
         Left            =   4380
         TabIndex        =   58
         Top             =   465
         Width           =   1410
      End
      Begin VB.CheckBox chkEco3 
         Caption         =   "Gestión"
         Height          =   240
         Left            =   4380
         TabIndex        =   57
         Top             =   780
         Width           =   1410
      End
      Begin VB.TextBox txtParticipacion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3390
         MaxLength       =   5
         TabIndex        =   25
         Top             =   810
         Width           =   870
      End
      Begin VB.CommandButton cmdGrabarG3 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3330
         TabIndex        =   26
         Top             =   1200
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancelarG3 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4575
         TabIndex        =   27
         Top             =   1200
         Width           =   1185
      End
      Begin VB.ComboBox cboG3 
         Height          =   315
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   810
         Width           =   2160
      End
      Begin VB.TextBox txtPersNombreG3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         TabIndex        =   23
         Top             =   435
         Width           =   4155
      End
      Begin Sicmact.TxtBuscar txtPersCodG3 
         Height          =   330
         Left            =   2340
         TabIndex        =   22
         Top             =   45
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCodigo6 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   1455
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label lblCodigo5 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   1200
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label lblCodigo4 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   900
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Partic."
         Height          =   195
         Left            =   2895
         TabIndex        =   42
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
         Height          =   195
         Left            =   105
         TabIndex        =   41
         Top             =   810
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   210
         TabIndex        =   40
         Top             =   135
         Width           =   555
      End
   End
   Begin VB.PictureBox PG3A 
      Height          =   2130
      Left            =   75
      ScaleHeight     =   2070
      ScaleWidth      =   10320
      TabIndex        =   37
      Top             =   4125
      Width           =   10380
      Begin VB.CommandButton cmdModificarG3 
         Caption         =   "&Modificar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7800
         TabIndex        =   10
         Top             =   1680
         Width           =   1185
      End
      Begin VB.CommandButton cmdNuevoG3 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6540
         TabIndex        =   9
         Top             =   1680
         Width           =   1185
      End
      Begin VB.CommandButton cmdEliminarG3 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9045
         TabIndex        =   11
         Top             =   1680
         Width           =   1185
      End
      Begin MSDataGridLib.DataGrid DG3 
         Height          =   1365
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   2408
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "cGECod"
            Caption         =   "Grupo"
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
            DataField       =   "cPersCodRel"
            Caption         =   "Codigo Rel"
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
            DataField       =   "cPersCodVinc"
            Caption         =   "Persona"
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
            DataField       =   "cPersNombre1"
            Caption         =   "Nombre"
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
            DataField       =   "nCargo"
            Caption         =   "TipoCargo"
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
            DataField       =   "cConsDescripcion2"
            Caption         =   "Cargo"
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
         BeginProperty Column06 
            DataField       =   "nParticip"
            Caption         =   "Participación"
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
         BeginProperty Column07 
            DataField       =   "nRela1"
            Caption         =   "nRela1"
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
         BeginProperty Column08 
            DataField       =   "nRela2"
            Caption         =   "nRela2"
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
         BeginProperty Column09 
            DataField       =   "nRela3"
            Caption         =   "nRela3"
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
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2654.929
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblCantidad3 
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
         Height          =   285
         Left            =   75
         TabIndex        =   45
         Top             =   1710
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Relación de Grupo Económico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   38
         Top             =   15
         Width           =   2475
      End
   End
   Begin VB.PictureBox PG2A 
      Height          =   2115
      Left            =   90
      ScaleHeight     =   2055
      ScaleWidth      =   10320
      TabIndex        =   32
      Top             =   1890
      Width           =   10380
      Begin VB.CommandButton cmdModificarG2 
         Caption         =   "&Modificar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7800
         TabIndex        =   6
         Top             =   1680
         Width           =   1185
      End
      Begin VB.CommandButton cmdEliminarG2 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9045
         TabIndex        =   7
         Top             =   1680
         Width           =   1185
      End
      Begin VB.CommandButton cmdNuevoG2 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6540
         TabIndex        =   5
         Top             =   1680
         Width           =   1185
      End
      Begin MSDataGridLib.DataGrid DG2 
         Height          =   1365
         Left            =   60
         TabIndex        =   4
         Top             =   240
         Width           =   10170
         _ExtentX        =   17939
         _ExtentY        =   2408
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "cGECod"
            Caption         =   "Grupo"
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
            DataField       =   "cPersCodRel"
            Caption         =   "Codigo"
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
            DataField       =   "cPersNombre"
            Caption         =   "Nombre"
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
            DataField       =   "nPrdPersRelac"
            Caption         =   "Tipo"
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
            DataField       =   "cConsDescripcion1"
            Caption         =   "Relación"
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
            DataField       =   "cTexto"
            Caption         =   "Control"
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
         BeginProperty Column06 
            DataField       =   "nRela1"
            Caption         =   "nRela1"
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
         BeginProperty Column07 
            DataField       =   "nRela2"
            Caption         =   "nRela2"
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
         BeginProperty Column08 
            DataField       =   "nRela3"
            Caption         =   "nRela3"
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
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1409.953
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4275.213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1604.976
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2489.953
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblCantidad2 
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
         Height          =   285
         Left            =   75
         TabIndex        =   44
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Grupo Económico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   33
         Top             =   15
         Width           =   1455
      End
   End
   Begin VB.PictureBox PG1A 
      Height          =   1755
      Left            =   90
      ScaleHeight     =   1695
      ScaleWidth      =   10305
      TabIndex        =   31
      Top             =   75
      Width           =   10365
      Begin VB.CommandButton cmdModificarG1 
         Caption         =   "&Modificar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7800
         TabIndex        =   2
         Top             =   1320
         Width           =   1185
      End
      Begin VB.CommandButton cmdEliminarG1 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9045
         TabIndex        =   3
         Top             =   1320
         Width           =   1185
      End
      Begin VB.CommandButton cmdNuevoG1 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6540
         TabIndex        =   1
         Top             =   1320
         Width           =   1185
      End
      Begin MSDataGridLib.DataGrid DG1 
         Height          =   1185
         Left            =   60
         TabIndex        =   0
         Top             =   75
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   2090
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "cGECod"
            Caption         =   "Codigo"
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
            DataField       =   "cGENom"
            Caption         =   "Nombre"
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
            DataField       =   "nGETipo"
            Caption         =   "Tipo"
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
            DataField       =   "cConsDescripcion"
            Caption         =   "Descripción"
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
            DataField       =   "cCodReporte"
            Caption         =   "CodRepo"
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
            DataField       =   "cDesReporte"
            Caption         =   "Reporte"
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
            MarqueeStyle    =   4
            BeginProperty Column00 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   120.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2835.213
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   120.189
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblCantidad1 
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
         Height          =   285
         Left            =   75
         TabIndex        =   43
         Top             =   1335
         Width           =   720
      End
   End
   Begin VB.PictureBox PG1B 
      Height          =   1635
      Left            =   60
      ScaleHeight     =   1575
      ScaleWidth      =   4380
      TabIndex        =   28
      Top             =   6420
      Visible         =   0   'False
      Width           =   4440
      Begin VB.ComboBox cboReporte 
         Height          =   315
         ItemData        =   "frmPersEcoGruRel1.frx":030A
         Left            =   915
         List            =   "frmPersEcoGruRel1.frx":031D
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   765
         Width           =   2010
      End
      Begin VB.CommandButton cmdCancelarG1 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3075
         TabIndex        =   16
         Top             =   1200
         Width           =   1185
      End
      Begin VB.CommandButton cmdGrabarG1 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1830
         TabIndex        =   15
         Top             =   1200
         Width           =   1185
      End
      Begin VB.ComboBox cboG1 
         Height          =   315
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   405
         Width           =   3315
      End
      Begin VB.TextBox txtNombre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   930
         MaxLength       =   40
         TabIndex        =   12
         Top             =   75
         Width           =   3315
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Reporte"
         Height          =   195
         Left            =   150
         TabIndex        =   52
         Top             =   780
         Width           =   570
      End
      Begin VB.Label lblCodigo1 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   120
         TabIndex        =   46
         Top             =   1215
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   135
         TabIndex        =   30
         Top             =   480
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   150
         TabIndex        =   29
         Top             =   120
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmPersEcoGruRel1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim banderaNuevoEdita As Integer

 
Private Sub cboG1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboReporte.SetFocus
End If
End Sub

Private Sub CBOG2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTexto.SetFocus
End If
End Sub

Private Sub cboG3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtParticipacion.SetFocus
End If
End Sub

Private Sub cboReporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabarG1.SetFocus
End If
End Sub

Private Sub cmdCancelarG1_Click()
    PG1B.Visible = False
    PG2B.Visible = False
    PG3B.Visible = False
    
    DG1.Enabled = True
    
    PG2A.Enabled = True
    PG3A.Enabled = True
    
    cmdNuevoG1.Enabled = True
    cmdModificarG1.Enabled = True
    cmdEliminarG1.Enabled = True
    
    cboReporte.Enabled = True
    
    banderaNuevoEdita = 0
    
    MuestraDG2
    If Val(lblCantidad2.Caption) > 0 Then
        MuestraDG3
        If Val(lblCantidad3.Caption) > 0 Then
            MuestraDG4
        Else
            BlanqueaDG4
        End If
    Else
        BlanqueaDG3
        BlanqueaDG4
    End If
    
    cmdNuevoG1.SetFocus
    
End Sub

Private Sub cmdCancelarG2_Click()
    PG1B.Visible = False
    PG2B.Visible = False
    PG3B.Visible = False
    
    DG2.Enabled = True
    
    PG1A.Enabled = True
    PG3A.Enabled = True
    
    cmdNuevoG2.Enabled = True
    cmdModificarG2.Enabled = True
    cmdEliminarG2.Enabled = True
    
    banderaNuevoEdita = 0
    
    MuestraDG3
    If Val(lblCantidad3.Caption) > 0 Then
        MuestraDG4
    Else
        BlanqueaDG4
    End If

    cmdNuevoG2.SetFocus

End Sub

Private Sub cmdCancelarG3_Click()
    PG1B.Visible = False
    PG2B.Visible = False
    PG3B.Visible = False
    
    DG3.Enabled = True
    
    PG1A.Enabled = True
    PG2A.Enabled = True
    
    banderaNuevoEdita = 0
    
    cmdNuevoG3.Enabled = True
    cmdModificarG3.Enabled = True
    cmdEliminarG3.Enabled = True
    
    If Val(lblCantidad3.Caption) > 0 Then
        MuestraDG4
    Else
        BlanqueaDG4
    End If
    
    cmdNuevoG3.SetFocus

End Sub

Private Sub cmdEliminarG1_Click()

    Dim rel As New DGrupoEco1
    
    On Error GoTo ErrorEliminar
    
    If Val(lblCantidad1.Caption) = 0 Then
        MsgBox "No hay registros que eliminar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Desea eliminar el registro seleccionado", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        If rel.getNumHijospersgrupoecon(Trim(lblCodigo1.Caption)) > 0 Then
            MsgBox "Ud. debe eliminar primero los registros relacionados", vbInformation, "Aviso"
            Exit Sub
        End If
        
        rel.EliminaPersGrupoEcon Trim(lblCodigo1.Caption)
        
    End If

    Set rel = Nothing

    MsgBox "Datos eliminados Satisfactoriamente", vbInformation, "Aviso "
    
    PG1B.Visible = False
    PG2B.Visible = False
    PG3B.Visible = False
    
    PG2A.Enabled = True
    PG3A.Enabled = True
    
    cmdNuevoG1.Enabled = True
    cmdModificarG1.Enabled = True
    cmdEliminarG1.Enabled = True
    
    LlenaDG1
    
    MuestraDG2
    If Val(lblCantidad2.Caption) > 0 Then
        MuestraDG3
        If Val(lblCantidad3.Caption) > 0 Then
            MuestraDG4
        Else
            BlanqueaDG4
        End If
    Else
        BlanqueaDG3
        BlanqueaDG4
    End If
    
    Exit Sub
    
ErrorEliminar:
    MsgBox "No se pudo efectuar grabación", vbExclamation, "Aviso"


End Sub

Private Sub cmdEliminarG2_Click()
    Dim rel As New DGrupoEco1
    
    On Error GoTo ErrorEliminar
    
    If Val(lblCantidad2.Caption) = 0 Then
        MsgBox "No hay registros que eliminar", vbInformation, "Aviso"
        Exit Sub
    End If

    If MsgBox("Desea eliminar el registro seleccionado", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        If rel.getNumHijosPersGEPersVinc(Trim(lblCodigo2.Caption), Trim(lblCodigo3.Caption)) > 0 Then
            MsgBox "Ud. debe eliminar primero los registros relacionados", vbInformation, "Aviso"
            Exit Sub
        End If
        
        rel.EliminaPersGERelacion Trim(lblCodigo2.Caption), Trim(lblCodigo3.Caption)
        
    End If

    Set rel = Nothing

    MsgBox "Datos eliminados Satisfactoriamente", vbInformation, "Aviso "
    
    PG1B.Visible = False
    PG2B.Visible = False
    PG3B.Visible = False
    
    PG2A.Enabled = True
    PG3A.Enabled = True
    
    cmdNuevoG1.Enabled = True
    cmdModificarG1.Enabled = True
    cmdEliminarG1.Enabled = True
    
    MuestraDG2
    If Val(lblCantidad2.Caption) > 0 Then
        MuestraDG3
        If Val(lblCantidad3.Caption) > 0 Then
            MuestraDG4
        Else
            BlanqueaDG4
        End If
    Else
        BlanqueaDG3
        BlanqueaDG4
    End If
    
    Exit Sub
    
ErrorEliminar:
    MsgBox "No se pudo efectuar grabación", vbExclamation, "Aviso"

End Sub

Private Sub cmdEliminarG3_Click()
    Dim rel As New DGrupoEco1
    
    On Error GoTo ErrorEliminar
    
    If Val(lblCantidad3.Caption) = 0 Then
        MsgBox "No hay registros que eliminar", vbInformation, "Aviso"
        Exit Sub
    End If

    If MsgBox("Desea eliminar el registro seleccionado", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        
        rel.EliminaPersGEPersVinc Trim(lblCodigo4.Caption), Trim(lblCodigo5.Caption), Trim(lblCodigo6.Caption)
        
    End If

    Set rel = Nothing

    MsgBox "Datos eliminados Satisfactoriamente", vbInformation, "Aviso "
    
    PG1B.Visible = False
    PG2B.Visible = False
    PG3B.Visible = False
    
    PG2A.Enabled = True
    PG3A.Enabled = True
    
    cmdNuevoG1.Enabled = True
    cmdModificarG1.Enabled = True
    cmdEliminarG1.Enabled = True
    
    MuestraDG2
    If Val(lblCantidad2.Caption) > 0 Then
        MuestraDG3
        If Val(lblCantidad3.Caption) > 0 Then
            MuestraDG4
        Else
            BlanqueaDG4
        End If
    Else
        BlanqueaDG3
        BlanqueaDG4
    End If
    
    Exit Sub
    
ErrorEliminar:
    MsgBox "No se pudo efectuar grabación", vbExclamation, "Aviso"

End Sub

Private Sub cmdGrabarG1_Click()
    
    Dim rel As New DGrupoEco1
    
    On Error GoTo ErrorGrabar
    
    If Len(Trim(txtNombre.Text)) = 0 Then
        MsgBox "Ud. Debe Ingresar un Nombre", vbExclamation, "Aviso"
        txtNombre.SetFocus
        Exit Sub
    End If
    If Len(Trim(cboG1.Text)) = 0 Then
        MsgBox "Ud. debe Seleccionar un Tipo", vbInformation, "Aviso"
        cboG1.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(cboReporte.Text)) = 0 Then
        MsgBox "Ud. debe Seleccionar un Reporte", vbInformation, "Aviso"
        cboReporte.SetFocus
        Exit Sub
    End If
    
    If banderaNuevoEdita = 1 Then
        If rel.GetExisteReporte(Trim(Right(cboReporte.Text, 3))) = 1 Then
            MsgBox "Relación de Reporte ya está asignado", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    rel.GrabaPersGrupoEcon Trim(lblCodigo1.Caption), txtNombre.Text, Val(Right(cboG1.Text, 4)), Trim(Right(cboReporte.Text, 3)), banderaNuevoEdita
    Set rel = Nothing
    
    banderaNuevoEdita = 0
    
    MsgBox "Datos Grabados Satisfactoriamente", vbInformation, "Aviso "
    
    PG1B.Visible = False
    PG2B.Visible = False
    PG3B.Visible = False
    
    DG1.Enabled = True
    
    PG2A.Enabled = True
    PG3A.Enabled = True
    
    cmdNuevoG1.Enabled = True
    cmdModificarG1.Enabled = True
    cmdEliminarG1.Enabled = True
    
    cboReporte.Enabled = True
    
    LlenaDG1
    
    MuestraDG2
    If Val(lblCantidad2.Caption) > 0 Then
        MuestraDG3
        If Val(lblCantidad3.Caption) > 0 Then
            MuestraDG4
        Else
            BlanqueaDG4
        End If
    Else
        BlanqueaDG3
        BlanqueaDG4
    End If
    
    cmdNuevoG1.SetFocus
    
    Exit Sub
    
ErrorGrabar:
    MsgBox "No se pudo efectuar grabación", vbExclamation, "Aviso"
    
    
End Sub

Private Sub cmdGrabarG2_Click()
    
    Dim rel As New DGrupoEco1
    
    On Error GoTo ErrorGrabar
    
    If Len(Trim(txtPersNombreG2.Text)) = 0 Then
        MsgBox "Ud. Debe Ingresar un Nombre", vbExclamation, "Aviso"
        txtPersCodG2.SetFocus
        Exit Sub
    End If
    If Len(Trim(CBOG2.Text)) = 0 Then
        MsgBox "Ud. debe Seleccionar una relación", vbInformation, "Aviso"
        CBOG2.SetFocus
        Exit Sub
    End If
        
    If banderaNuevoEdita = 1 Then
        If rel.getExistePersonaEnPersGERelacion(Trim(lblCodigo2.Caption), Trim(txtPersCodG2.Text)) = 1 Then
            MsgBox "Relación ya está registrada", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    If banderaNuevoEdita = 1 Then
        rel.GrabaPersGeRelacion Trim(lblCodigo1.Caption), "", Trim(txtPersCodG2.Text), Val(Right(CBOG2.Text, 4)), Trim(txtTexto.Text), chkEcoa.value, chkEcob.value, chkEcoc.value, banderaNuevoEdita
    ElseIf banderaNuevoEdita = 2 Then
        rel.GrabaPersGeRelacion Trim(lblCodigo2.Caption), Trim(lblCodigo3.Caption), Trim(txtPersCodG2.Text), Val(Right(CBOG2.Text, 4)), Trim(txtTexto.Text), chkEcoa.value, chkEcob.value, chkEcoc.value, banderaNuevoEdita
    End If
    
    Set rel = Nothing
    
    banderaNuevoEdita = 0
    
    MsgBox "Datos Grabados Satisfactoriamente", vbInformation, "Aviso "
    
    
    PG1B.Visible = False
    PG2B.Visible = False
    PG3B.Visible = False
    
    DG2.Enabled = True
    
    PG1A.Enabled = True
    PG3A.Enabled = True
    
    cmdNuevoG2.Enabled = True
    cmdModificarG2.Enabled = True
    cmdEliminarG2.Enabled = True
    
    
    MuestraDG2
    If Val(lblCantidad2.Caption) > 0 Then
        MuestraDG3
        If Val(lblCantidad3.Caption) > 0 Then
            MuestraDG4
        Else
            BlanqueaDG4
        End If
    Else
        BlanqueaDG3
        BlanqueaDG4
    End If

    cmdNuevoG2.SetFocus

    Exit Sub
    
ErrorGrabar:
    MsgBox "No se pudo efectuar grabación", vbExclamation, "Aviso"

End Sub

Private Sub cmdGrabarG3_Click()
    
    Dim rel As New DGrupoEco1
    
    On Error GoTo ErrorGrabar
    
    If Len(Trim(txtPersNombreG3.Text)) = 0 Then
        MsgBox "Ud. Debe Ingresar un Nombre", vbExclamation, "Aviso"
        txtPersNombreG3.SetFocus
        Exit Sub
    End If
    If Len(Trim(cboG3.Text)) = 0 Then
        MsgBox "Ud. debe Seleccionar un cargo", vbInformation, "Aviso"
        cboG3.SetFocus
        Exit Sub
    End If
    
    If banderaNuevoEdita = 1 Then
        rel.GrabaPersGEPersVinc Trim(lblCodigo2.Caption), Trim(lblCodigo3.Caption), "", Trim(txtPersCodG3.Text), Val(Right(cboG3.Text, 4)), Val(txtParticipacion.Text), chkEco1.value, chkEco2.value, chkEco3.value, banderaNuevoEdita
    ElseIf banderaNuevoEdita = 2 Then
        rel.GrabaPersGEPersVinc Trim(lblCodigo4.Caption), Trim(lblCodigo5.Caption), Trim(lblCodigo6.Caption), Trim(txtPersCodG3.Text), Val(Right(cboG3.Text, 4)), Val(txtParticipacion.Text), chkEco1.value, chkEco2.value, chkEco3.value, banderaNuevoEdita
    End If
    
    Set rel = Nothing
    
    banderaNuevoEdita = 0
    
    MsgBox "Datos Grabados Satisfactoriamente", vbInformation, "Aviso "
    
    
    PG1B.Visible = False
    PG2B.Visible = False
    PG3B.Visible = False
    
    DG3.Enabled = True
    
    PG1A.Enabled = True
    PG2A.Enabled = True
    
    cmdNuevoG3.Enabled = True
    cmdModificarG3.Enabled = True
    cmdEliminarG3.Enabled = True
    
    If Val(lblCantidad2.Caption) > 0 Then
        MuestraDG3
        If Val(lblCantidad3.Caption) > 0 Then
            MuestraDG4
        Else
            BlanqueaDG4
        End If
    Else
        BlanqueaDG3
        BlanqueaDG4
    End If
    
    cmdNuevoG3.SetFocus
    
    Exit Sub
    
ErrorGrabar:
    MsgBox "No se pudo efectuar grabación", vbExclamation, "Aviso"
    
    
End Sub

Private Sub cmdModificarG1_Click()

    If Val(lblCantidad1.Caption) = 0 Then
        MsgBox "No hay registros que modificar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    banderaNuevoEdita = 2
    
    PG1B.Visible = True
    PG2B.Visible = False
    PG3B.Visible = False
    
    DG1.Enabled = False
    
    PG2A.Enabled = False
    PG3A.Enabled = False
    
    cmdNuevoG1.Enabled = False
    cmdModificarG1.Enabled = False
    cmdEliminarG1.Enabled = False

    cboReporte.Enabled = False

End Sub

Private Sub cmdModificarG2_Click()
    
If Val(lblCantidad2.Caption) = 0 Then
    MsgBox "No hay registros que modificar", vbInformation, "Aviso"
    Exit Sub
End If

    banderaNuevoEdita = 2
    
    PG1B.Visible = False
    PG2B.Visible = True
    PG3B.Visible = False
    
    DG2.Enabled = False
    
    PG1A.Enabled = False
    PG3A.Enabled = False
    
    cmdNuevoG2.Enabled = False
    cmdModificarG2.Enabled = False
    cmdEliminarG2.Enabled = False
End Sub

Private Sub cmdModificarG3_Click()
    
    If Val(lblCantidad3.Caption) = 0 Then
        MsgBox "No hay registros que modificar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    
    banderaNuevoEdita = 2
    
    PG1B.Visible = False
    PG2B.Visible = False
    PG3B.Visible = True
    
    DG3.Enabled = False
    
    PG1A.Enabled = False
    PG2A.Enabled = False
    
    cmdNuevoG3.Enabled = False
    cmdModificarG3.Enabled = False
    cmdEliminarG3.Enabled = False

End Sub

Private Sub cmdNuevoG1_Click()
    
    banderaNuevoEdita = 1
    
    PG1B.Visible = True
    PG2B.Visible = False
    PG3B.Visible = False
        
    DG1.Enabled = False
    
    PG2A.Enabled = False
    PG3A.Enabled = False
    
    cmdNuevoG1.Enabled = False
    cmdModificarG1.Enabled = False
    cmdEliminarG1.Enabled = False
    
    txtNombre.Text = ""
    cboG1.ListIndex = -1
    lblCodigo1.Caption = ""
    cboReporte.ListIndex = -1
    
End Sub

Private Sub cmdNuevoG2_Click()
    
    banderaNuevoEdita = 1

    PG1B.Visible = False
    PG2B.Visible = True
    PG3B.Visible = False
    
    DG2.Enabled = False
    
    PG1A.Enabled = False
    PG3A.Enabled = False
    
    cmdNuevoG2.Enabled = False
    cmdModificarG2.Enabled = False
    cmdEliminarG2.Enabled = False
    
    txtPersCodG2.Text = ""
    txtPersNombreG2.Text = ""
    'CBOG2.ListIndex = -1
    
    lblCodigo2.Caption = ""
    lblCodigo3.Caption = ""
    txtTexto.Text = ""
    
    chkEcoa.value = 0
    chkEcob.value = 0
    chkEcoc.value = 0
    
End Sub

Private Sub cmdNuevoG3_Click()
    
    banderaNuevoEdita = 1
    
    PG1B.Visible = False
    PG2B.Visible = False
    PG3B.Visible = True
    
    DG3.Enabled = False
    
    PG1A.Enabled = False
    PG2A.Enabled = False
    
    cmdNuevoG3.Enabled = False
    cmdModificarG3.Enabled = False
    cmdEliminarG3.Enabled = False
    
    txtPersCodG3.Text = ""
    txtPersNombreG3.Text = ""
    cboG3.ListIndex = -1
    txtParticipacion.Text = ""
    
    lblCodigo4.Caption = ""
    lblCodigo5.Caption = ""
    lblCodigo6.Caption = ""
    
    chkEco1.value = 0
    chkEco2.value = 0
    chkEco3.value = 0

End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub DG1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Val(lblCantidad1.Caption) > 0 Then
        MuestraDG2
        If Val(lblCantidad2.Caption) > 0 Then
            MuestraDG3
            If Val(lblCantidad3.Caption) > 0 Then
                MuestraDG4
            Else
                BlanqueaDG4
            End If
        Else
            BlanqueaDG3
            BlanqueaDG4
        End If
    Else
        BlanqueaDG2
        BlanqueaDG3
        BlanqueaDG4
    End If
End Sub

Private Sub MuestraDG2()
    LlenaDG2 DG1.Columns(0).Text
    txtNombre.Text = DG1.Columns(1).Text
    cboG1.ListIndex = IndiceListaCombo(cboG1, Val(DG1.Columns(2).Text))
    lblCodigo1.Caption = DG1.Columns(0).Text
    cboReporte.ListIndex = IndiceListaCombo(cboReporte, Trim(DG1.Columns(4).Text))


End Sub

Private Sub BlanqueaDG2()
    LlenaDG2 "xx"
    txtNombre.Text = ""
    cboG1.ListIndex = -1
    lblCodigo1.Caption = ""
    cboReporte.ListIndex = -1
End Sub

Private Sub DG2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Val(lblCantidad2.Caption) > 0 Then
        MuestraDG3
        If Val(lblCantidad3.Caption) > 0 Then
            MuestraDG4
        Else
            BlanqueaDG4
        End If
    Else
        BlanqueaDG3
        BlanqueaDG4
    End If
End Sub

Private Sub MuestraDG3()
    LlenaDG3 DG2.Columns(0).Text, DG2.Columns(1).Text
    txtPersCodG2.Text = DG2.Columns(1).Text
    txtPersNombreG2.Text = DG2.Columns(2).Text
    CBOG2.ListIndex = IndiceListaCombo(CBOG2, Val(DG2.Columns(3).Text))
    lblCodigo2.Caption = DG2.Columns(0).Text
    lblCodigo3.Caption = DG2.Columns(1).Text
    txtTexto.Text = DG2.Columns(5).Text

    chkEcoa.value = Val(DG2.Columns(6).Text)
    chkEcob.value = Val(DG2.Columns(7).Text)
    chkEcoc.value = Val(DG2.Columns(8).Text)


End Sub

Private Sub BlanqueaDG3()
    LlenaDG3 "xx", "xx"
    txtPersCodG2.Text = ""
    txtPersNombreG2.Text = ""
    CBOG2.ListIndex = -1
    lblCodigo2.Caption = ""
    lblCodigo3.Caption = ""
    txtTexto.Text = ""
    
    chkEcoa.value = 0
    chkEcob.value = 0
    chkEcoc.value = 0
    
End Sub

Private Sub DG3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Val(lblCantidad3.Caption) > 0 Then
        MuestraDG4
    Else
        BlanqueaDG4
    End If

End Sub

Private Sub MuestraDG4()
    txtPersCodG3.Text = DG3.Columns(2).Text
    txtPersNombreG3.Text = DG3.Columns(3).Text
    cboG3.ListIndex = IndiceListaCombo(cboG3, Val(DG3.Columns(4).Text))
    txtParticipacion.Text = DG3.Columns(6).Text
    lblCodigo4.Caption = DG3.Columns(0).Text
    lblCodigo5.Caption = DG3.Columns(1).Text
    lblCodigo6.Caption = DG3.Columns(2).Text
    
    chkEco1.value = Val(DG3.Columns(7).Text)
    chkEco2.value = Val(DG3.Columns(8).Text)
    chkEco3.value = Val(DG3.Columns(9).Text)
    
End Sub

Private Sub BlanqueaDG4()
    txtPersCodG3.Text = ""
    txtPersNombreG3.Text = ""
    cboG3.ListIndex = -1
    txtParticipacion.Text = ""
    lblCodigo4.Caption = ""
    lblCodigo5.Caption = ""
    lblCodigo6.Caption = ""
    
    chkEco1.value = 0
    chkEco2.value = 0
    chkEco3.value = 0
    
End Sub

Private Sub Form_Load()
    
    'Combo Nro 01
    LlenaComboConstante 4027, cboG1
    
    'Combo Nro 02
    LlenaComboConstante 4028, CBOG2
    
    'Combo Nro 03
    LlenaComboConstante 4029, cboG3
    
    'Persona G2
    txtPersCodG2.TipoBusPers = BusPersCodigo
    txtPersCodG2.TipoBusqueda = BuscaPersona
    txtPersCodG2.EditFlex = False
    
    'Persona G3
    txtPersCodG3.TipoBusPers = BusPersCodigo
    txtPersCodG3.TipoBusqueda = BuscaPersona
    txtPersCodG3.EditFlex = False
    
    LlenaDG1
    
    If DG1.ApproxCount <> 0 Then
        LlenaDG2 DG1.Columns(0).Text
        
        LlenaDG3 DG2.Columns(0).Text, DG2.Columns(1).Text
    
        DG1.Columns(2).Width = 0
        DG1.Columns(4).Width = 0
        
        DG2.Columns(0).Width = 0
        DG2.Columns(3).Width = 0
        DG2.Columns(4).Width = 0
        DG2.Columns(6).Width = 0
        DG2.Columns(7).Width = 0
        DG2.Columns(8).Width = 0
        
        DG3.Columns(0).Width = 0
        DG3.Columns(1).Width = 0
        'DG3.Columns(4).Width = 0
        DG3.Columns(7).Width = 0
        DG3.Columns(8).Width = 0
        DG3.Columns(9).Width = 0
        CentraForm Me
    End If
End Sub

Private Sub LlenaDG1()

Dim rel As New DGrupoEco1
Dim nCantidad As Long

    nCantidad = 0
    Set DG1.DataSource = rel.GetG1(nCantidad)
    lblCantidad1.Caption = nCantidad
    
Set rel = Nothing

End Sub

Private Sub LlenaDG2(cGECod As String)

Dim rel As New DGrupoEco1
Dim nCantidad As Long

    nCantidad = 0
    Set DG2.DataSource = rel.GetG2(nCantidad, cGECod)
    lblCantidad2.Caption = nCantidad

Set rel = Nothing

End Sub

Private Sub LlenaDG3(cGECod As String, cPersCodRel As String)

Dim rel As New DGrupoEco1
Dim nCantidad As Long

    nCantidad = 0
    Set DG3.DataSource = rel.GetG3(nCantidad, cGECod, cPersCodRel)
    lblCantidad3.Caption = nCantidad
    
Set rel = Nothing

End Sub
 

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboG1.SetFocus
    End If
End Sub

Private Sub txtParticipacion_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmdGrabarG3.SetFocus
    End If
End Sub

Private Sub txtPersCodG2_EmiteDatos()
txtPersNombreG2 = txtPersCodG2.psDescripcion
If txtPersNombreG2 <> "" Then
    CBOG2.SetFocus
End If
End Sub

Private Sub txtPersCodG3_EmiteDatos()
txtPersNombreG3 = txtPersCodG3.psDescripcion
If txtPersNombreG3 <> "" Then
    cboG3.SetFocus
End If
End Sub

Private Sub txtTexto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabarG2.SetFocus
    End If

End Sub
