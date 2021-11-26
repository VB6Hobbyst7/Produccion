VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredNivAprobacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Niveles de Aprobacion de Credito"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   Icon            =   "frmCredNivAprobacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   390
      Left            =   7110
      TabIndex        =   47
      Top             =   7140
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6915
      Left            =   30
      TabIndex        =   0
      Top             =   150
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   12197
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Cargos de Creditos"
      TabPicture(0)   =   "frmCredNivAprobacion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Rangos de Aprobacion"
      TabPicture(1)   =   "frmCredNivAprobacion.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame7 
         Caption         =   "Definicion de Niveles"
         Height          =   2400
         Left            =   -74895
         TabIndex        =   38
         Top             =   3720
         Width           =   8400
         Begin MSDataGridLib.DataGrid DGNiveles 
            Height          =   1965
            Left            =   105
            TabIndex        =   39
            Top             =   240
            Width           =   6675
            _ExtentX        =   11774
            _ExtentY        =   3466
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
               DataField       =   "nItem"
               Caption         =   "Item"
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
               DataField       =   "cCodRango"
               Caption         =   "Rango"
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
               DataField       =   "cCodCargo"
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
            BeginProperty Column03 
               DataField       =   "cNomCargo"
               Caption         =   "Descripcion de Cargo"
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
               DataField       =   "cComen"
               Caption         =   "Comentario"
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
                  ColumnWidth     =   494.929
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   810.142
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   764.787
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2984.882
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   929.764
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame8 
            Height          =   1020
            Left            =   135
            TabIndex        =   42
            Top             =   1080
            Width           =   6615
            Begin VB.CheckBox ChkComen 
               Caption         =   "Comentario Obligartorio"
               Height          =   255
               Left            =   855
               TabIndex        =   48
               Top             =   645
               Width           =   2895
            End
            Begin VB.CommandButton CmdNivCancelar 
               Caption         =   "Cancelar"
               Height          =   315
               Left            =   5055
               TabIndex        =   46
               Top             =   585
               Width           =   1335
            End
            Begin VB.CommandButton CmdNivAceptar 
               Caption         =   "Aceptar"
               Height          =   315
               Left            =   5055
               TabIndex        =   45
               Top             =   210
               Width           =   1335
            End
            Begin VB.ComboBox CboCargo 
               Height          =   315
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Top             =   240
               Width           =   3765
            End
            Begin VB.Label Label8 
               Caption         =   "Cargo  :"
               Height          =   240
               Left            =   150
               TabIndex        =   43
               Top             =   270
               Width           =   615
            End
         End
         Begin VB.CommandButton CmdNivElim 
            Caption         =   "Eliminar"
            Height          =   390
            Left            =   7005
            TabIndex        =   41
            Top             =   720
            Width           =   1200
         End
         Begin VB.CommandButton CmdNivNuevo 
            Caption         =   "Nuevo"
            Height          =   390
            Left            =   7005
            TabIndex        =   40
            Top             =   240
            Width           =   1200
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3330
         Left            =   -74895
         TabIndex        =   21
         Top             =   375
         Width           =   8385
         Begin MSDataGridLib.DataGrid DGRangos 
            Height          =   2940
            Left            =   90
            TabIndex        =   22
            Top             =   270
            Width           =   6870
            _ExtentX        =   12118
            _ExtentY        =   5186
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
               DataField       =   "cCodRango"
               Caption         =   "Rango"
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
               DataField       =   "cConsDescripcion"
               Caption         =   "Producto"
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
               DataField       =   "cMonedaDesc"
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
            BeginProperty Column03 
               DataField       =   "cTipoCredDesc"
               Caption         =   "Tipo de Cred."
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
               DataField       =   "nMontoMin"
               Caption         =   "Monto Min."
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
               DataField       =   "nMontoMax"
               Caption         =   "Monto Max."
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
                  ColumnWidth     =   2009.764
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   915.024
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1019.906
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton CmdRangElim 
            Caption         =   "Eliminar"
            Height          =   345
            Left            =   7140
            TabIndex        =   37
            Top             =   750
            Width           =   1110
         End
         Begin VB.CommandButton CmdRangNuevo 
            Caption         =   "Nuevo"
            Height          =   345
            Left            =   7140
            TabIndex        =   36
            Top             =   315
            Width           =   1110
         End
         Begin VB.Frame Frame6 
            Height          =   1995
            Left            =   90
            TabIndex        =   23
            Top             =   1200
            Width           =   6825
            Begin VB.CommandButton CmdRangCancel 
               Caption         =   "Cancelar"
               Height          =   345
               Left            =   5445
               TabIndex        =   35
               Top             =   1515
               Width           =   1110
            End
            Begin VB.CommandButton CmdRangAceptar 
               Caption         =   "Aceptar"
               Height          =   345
               Left            =   4290
               TabIndex        =   34
               Top             =   1515
               Width           =   1110
            End
            Begin VB.TextBox TxtMontoMin 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1350
               TabIndex        =   33
               Text            =   "0.00"
               Top             =   1080
               Width           =   1245
            End
            Begin VB.TextBox TxtMontoMax 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5085
               TabIndex        =   31
               Text            =   "0.00"
               Top             =   1080
               Width           =   1440
            End
            Begin VB.ComboBox CboTipoCred 
               Height          =   315
               Left            =   1350
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   660
               Width           =   2445
            End
            Begin VB.ComboBox CboMoneda 
               Height          =   315
               Left            =   4680
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   225
               Width           =   1845
            End
            Begin VB.ComboBox CboProd 
               Height          =   315
               Left            =   1005
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   210
               Width           =   2775
            End
            Begin VB.Label Label7 
               Caption         =   "Monto Minimo  :"
               Height          =   225
               Left            =   105
               TabIndex        =   32
               Top             =   1125
               Width           =   1230
            End
            Begin VB.Label Label6 
               Caption         =   "Monto Maximo  :"
               Height          =   225
               Left            =   3855
               TabIndex        =   30
               Top             =   1125
               Width           =   1230
            End
            Begin VB.Label Label5 
               Caption         =   "Tipo de Credito :"
               Height          =   195
               Left            =   105
               TabIndex        =   28
               Top             =   705
               Width           =   1275
            End
            Begin VB.Label Label4 
               Caption         =   "Moneda  :"
               Height          =   195
               Left            =   3855
               TabIndex        =   26
               Top             =   270
               Width           =   855
            End
            Begin VB.Label Label3 
               Caption         =   "Producto   :"
               Height          =   225
               Left            =   105
               TabIndex        =   24
               Top             =   255
               Width           =   900
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Responsables de Cargo"
         Height          =   3105
         Left            =   135
         TabIndex        =   9
         Top             =   3660
         Width           =   8340
         Begin MSDataGridLib.DataGrid DGUsuCargo 
            Height          =   2760
            Left            =   90
            TabIndex        =   10
            Top             =   270
            Width           =   6690
            _ExtentX        =   11800
            _ExtentY        =   4868
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "cCodCargo"
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
            BeginProperty Column01 
               DataField       =   "cCodUsu"
               Caption         =   "Usuario"
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
               DataField       =   "cNomUsu"
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
               DataField       =   "cAgeCod"
               Caption         =   "Cod Age"
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
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1065.26
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   3404.977
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   794.835
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame4 
            Height          =   1440
            Left            =   75
            TabIndex        =   13
            Top             =   1605
            Width           =   6645
            Begin VB.ComboBox cboAgencia 
               Height          =   315
               Left            =   870
               Style           =   2  'Dropdown List
               TabIndex        =   51
               Top             =   1013
               Width           =   3990
            End
            Begin VB.CommandButton CmdUsuCancelar 
               Caption         =   "Cancelar"
               Height          =   330
               Left            =   5325
               TabIndex        =   20
               Top             =   615
               Width           =   1095
            End
            Begin VB.CommandButton CmdUsuAceptar 
               Caption         =   "Aceptar"
               Height          =   345
               Left            =   5325
               TabIndex        =   19
               Top             =   225
               Width           =   1095
            End
            Begin VB.CommandButton CmdUsuCompNom 
               Caption         =   "Comprobar Nombre"
               Height          =   300
               Left            =   1755
               TabIndex        =   18
               Top             =   210
               Width           =   1620
            End
            Begin VB.TextBox TxtUsu 
               Alignment       =   2  'Center
               Height          =   300
               Left            =   870
               TabIndex        =   14
               Top             =   210
               Width           =   705
            End
            Begin VB.Label Label9 
               Caption         =   "Agencia:"
               Height          =   240
               Left            =   105
               TabIndex        =   50
               Top             =   1050
               Width           =   690
            End
            Begin VB.Label LblNomUsu 
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
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   870
               TabIndex        =   17
               Top             =   630
               Width           =   3990
            End
            Begin VB.Label Label2 
               Caption         =   "Nombre :"
               Height          =   270
               Left            =   105
               TabIndex        =   16
               Top             =   615
               Width           =   720
            End
            Begin VB.Label Label1 
               Caption         =   "Usuario :"
               Height          =   270
               Left            =   105
               TabIndex        =   15
               Top             =   225
               Width           =   735
            End
         End
         Begin VB.CommandButton CmdUsuEliminar 
            Caption         =   "Eliminar"
            Height          =   390
            Left            =   6930
            TabIndex        =   12
            Top             =   870
            Width           =   1305
         End
         Begin VB.CommandButton CmdUsuNuevo 
            Caption         =   "Nuevo"
            Height          =   390
            Left            =   6930
            TabIndex        =   11
            Top             =   405
            Width           =   1305
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3285
         Left            =   135
         TabIndex        =   1
         Top             =   360
         Width           =   8355
         Begin MSDataGridLib.DataGrid DGCargos 
            Height          =   2865
            Left            =   90
            TabIndex        =   2
            Top             =   255
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   5054
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "cCodCargo"
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
               DataField       =   "cNomCargo"
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
            BeginProperty Column02 
               DataField       =   "bTodaAgencia"
               Caption         =   "Para toda la Institucion"
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
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               BeginProperty Column00 
                  ColumnWidth     =   945.071
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   3135.118
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1725.165
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame2 
            Height          =   990
            Left            =   90
            TabIndex        =   5
            Top             =   2130
            Width           =   6060
            Begin VB.CheckBox chkTodasAgencias 
               Caption         =   "Para toda la Institución"
               Height          =   240
               Left            =   75
               TabIndex        =   49
               Top             =   600
               Width           =   2565
            End
            Begin VB.CommandButton CmdCrgCancelar 
               Caption         =   "Cancelar"
               Height          =   285
               Left            =   4845
               TabIndex        =   8
               Top             =   600
               Width           =   1050
            End
            Begin VB.CommandButton CmdCrgAceptar 
               Caption         =   "&Aceptar"
               Height          =   285
               Left            =   3765
               TabIndex        =   7
               Top             =   600
               Width           =   1050
            End
            Begin VB.TextBox TxtCargo 
               Height          =   315
               Left            =   75
               TabIndex        =   6
               Top             =   210
               Width           =   5850
            End
         End
         Begin VB.CommandButton CmdCrgEliminar 
            Caption         =   "&Eliminar"
            Height          =   420
            Left            =   6750
            TabIndex        =   4
            Top             =   810
            Width           =   1380
         End
         Begin VB.CommandButton CmdCrgNuevo 
            Caption         =   "&Nuevo"
            Height          =   420
            Left            =   6750
            TabIndex        =   3
            Top             =   300
            Width           =   1380
         End
      End
   End
End
Attribute VB_Name = "frmCredNivAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RCredCargos As New ADODB.Recordset
Dim RCredCargosResp As New ADODB.Recordset
Dim RCredRangos As New ADODB.Recordset
Dim RCredNiveles As New ADODB.Recordset

Private Sub HabilitaNiveles(ByVal pbHabilita As Boolean)

    If pbHabilita Then
        DGNiveles.Height = 765
        CboCargo.ListIndex = -1
    Else
        DGNiveles.Height = 1965
    End If
    CmdNivNuevo.Enabled = Not pbHabilita
    CmdNivElim.Enabled = Not pbHabilita
    DGNiveles.Enabled = Not pbHabilita

End Sub

Private Sub HabilitaRangos(ByVal pbHabilita As Boolean)

    If pbHabilita Then
        DGRangos.Height = 870
        CboProd.ListIndex = -1
        CboTipoCred.ListIndex = -1
        cboMoneda.ListIndex = -1
        txtMontoMin.Text = "0.00"
        txtMontoMax.Text = "0.00"
    Else
        DGRangos.Height = 2940
    End If
    
    DGRangos.Enabled = Not pbHabilita

End Sub

Private Sub CargaControles()
Dim D As COMDCredito.DCOMCredito
Dim TC() As String
Dim M() As String
Dim P() As String
Dim c() As String
Dim i As Integer
    Set D = New COMDCredito.DCOMCredito
    Call D.CargaCombosCredNivAprob(P, TC, M, c)
    Set D = Nothing

    CboProd.Clear
    For i = 0 To UBound(P) - 1
        CboProd.AddItem P(i)
    Next i

    Me.CboTipoCred.Clear
    For i = 0 To UBound(TC) - 1
        CboTipoCred.AddItem TC(i)
    Next i
    
    cboMoneda.Clear
    For i = 0 To UBound(M) - 1
        cboMoneda.AddItem M(i)
    Next i

    CboCargo.Clear
    For i = 0 To UBound(c) - 1
        CboCargo.AddItem c(i)
    Next i

    '06-05-2005
    Dim oCons As COMDConstantes.DCOMConstantes
    Set oCons = New COMDConstantes.DCOMConstantes
    Dim rsAge As ADODB.Recordset
    Set rsAge = oCons.GetAgencias
    
    Do While Not rsAge.EOF
        CboAgencia.AddItem Trim(rsAge!Columna1) & Space(100) & Trim(Str(rsAge!columna2))
        rsAge.MoveNext
    Loop
    rsAge.Close

    Set oCons = Nothing
End Sub

Private Sub CargaDatosRangos()
Dim D As COMDCredito.DCOMCredito

    Set D = New COMDCredito.DCOMCredito
    Set RCredRangos = D.RecuperaCredRangos
    Set D = Nothing
    
    Set DGRangos.DataSource = RCredRangos
    DGRangos.Refresh
    
End Sub

Private Sub CargaDatosNiveles(ByVal psCodRango As String)
Dim D As COMDCredito.DCOMCredito

    Set D = New COMDCredito.DCOMCredito
    Set RCredNiveles = D.RecuperaCredNiveles(psCodRango)
    Set D = Nothing
    
    Set DGNiveles.DataSource = RCredNiveles
    DGNiveles.Refresh
    
End Sub

Private Sub CargaDatosCargos()
Dim D As COMDCredito.DCOMCredito

    Set D = New COMDCredito.DCOMCredito
    Set RCredCargos = D.RecuperaCredCargos
    Set D = Nothing
    
    Set DGCargos.DataSource = RCredCargos
    DGCargos.Refresh
    
End Sub

Private Sub CargaDatosCargosResponsable(ByVal psCodCargo As String)
Dim D As COMDCredito.DCOMCredito

    Set D = New COMDCredito.DCOMCredito
    Set RCredCargosResp = D.RecuperaCredCargosResponsable(psCodCargo)
    Set D = Nothing
    
    Set DGUsuCargo.DataSource = RCredCargosResp
    DGUsuCargo.Refresh
    
End Sub

Private Sub CboCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdNivAceptar.SetFocus
    End If
End Sub

Private Sub CboMoneda_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
        CboTipoCred.SetFocus
    End If

End Sub

Private Sub CboProd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        cboMoneda.SetFocus
    End If
    
End Sub


Private Sub CboTipoCred_KeyPress(KeyAscii As Integer)
    
      If KeyAscii = 13 Then
        txtMontoMin.SetFocus
    End If
End Sub

Private Sub CmdCrgAceptar_Click()
Dim D As COMDCredito.DCOMCredito
Dim i As Integer

    If Trim(txtCargo.Text) = "" Then
        MsgBox "Ingrese el Nombre de Cargo"
        Exit Sub
    End If
    
    Set D = New COMDCredito.DCOMCredito
    Call D.IngresaCreCargo(txtCargo.Text, chkTodasAgencias.value)
    
    'ARCV 13-06-2006
    Dim R As ADODB.Recordset
    
    Set R = D.CargarCargosAprobacion
    CboCargo.Clear
    For i = 0 To R.RecordCount - 1
        CboCargo.AddItem R!cNomCargo & Space(100) & R!cCodCargo
        R.MoveNext
    Next i
    Set R = Nothing
    '---------
    Set D = Nothing
    
    Call CargaDatosCargos
    
    Call HabilitaIngresoCargo(False)
    
End Sub

Private Sub CmdCrgCancelar_Click()
    Call HabilitaIngresoCargo(False)
End Sub

Private Sub CmdCrgEliminar_Click()
Dim D As COMDCredito.DCOMCredito

    If RCredCargos.RecordCount = 0 Then
        MsgBox "No Existen Registros para Realizar esta Operacion", vbInformation, "Aviso"
    End If
    
    If MsgBox("Se van a Eliminar Registros, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set D = New COMDCredito.DCOMCredito
        Call D.EliminarCargo(RCredCargos!cCodCargo)
        Set D = Nothing
        Call CargaDatosCargos
    End If

End Sub

Private Sub CmdCrgNuevo_Click()

    Call HabilitaIngresoCargo(True)
End Sub

Private Sub CmdNivAceptar_Click()
Dim D As COMDCredito.DCOMCredito

    If CboCargo.ListIndex = -1 Then
        MsgBox "Seleccione el Nivel", vbInformation, "Aviso"
        CboCargo.SetFocus
        Exit Sub
    End If
    
    'ARCV 11-07-2006
    If ValidaCargo_x_Rango(RCredRangos!cCodRango, Trim(Right(CboCargo.Text, 3))) = False Then
        MsgBox "El cargo ya se encuentra registrado", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    Set D = New COMDCredito.DCOMCredito
    Call D.InsertaNiveles(RCredRangos!cCodRango, Trim(Right(CboCargo.Text, 3)), IIf(ChkComen.value = 1, True, False))
    Set D = Nothing
    
    Call CargaDatosNiveles(RCredRangos!cCodRango)
    
    HabilitaNiveles False
End Sub

Private Function ValidaCargo_x_Rango(ByVal psCodRango As String, _
                                    ByVal psCodCargo As String) As Boolean
ValidaCargo_x_Rango = True

If RCredNiveles.RecordCount = 0 Then
    ValidaCargo_x_Rango = True
    Exit Function
Else
    RCredNiveles.MoveFirst
End If

While Not RCredNiveles.EOF
    If RCredNiveles!cCodRango = psCodRango And RCredNiveles!cCodCargo = psCodCargo Then
        ValidaCargo_x_Rango = False
        Exit Function
    End If
    RCredNiveles.MoveNext
Wend
End Function

Private Sub CmdNivCancelar_Click()
    HabilitaNiveles False
End Sub

Private Sub CmdNivElim_Click()
Dim D As COMDCredito.DCOMCredito

    If RCredNiveles.RecordCount = 0 Then
        MsgBox "No Existen Registros para Realizar esta Operacion", vbInformation, "Aviso"
    End If
    
    If MsgBox("Se van a Eliminar Registros, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set D = New COMDCredito.DCOMCredito
        Call D.EliminarNivel(RCredNiveles!nItem, RCredNiveles!cCodRango, RCredNiveles!cCodCargo)
        Set D = Nothing
        Call CargaDatosNiveles(RCredRangos!cCodRango)
    End If

    Call HabilitaNiveles(False)
End Sub

Private Sub CmdNivNuevo_Click()

    HabilitaNiveles True
End Sub

Private Sub CmdRangAceptar_Click()
Dim D As COMDCredito.DCOMCredito

    If CboProd.ListIndex = -1 Then
        MsgBox "Seleccione un Producto", vbInformation, "Aviso"
        CboProd.SetFocus
    End If
    
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Seleccione una Moneda", vbInformation, "Aviso"
        cboMoneda.SetFocus
    End If
    
    If CboTipoCred.ListIndex = -1 Then
        MsgBox "Seleccione un Tipo de Credito", vbInformation, "Aviso"
        CboTipoCred.SetFocus
    End If
    
    If txtMontoMin.Text = "0.00" Then
        MsgBox "Digite un Monto Minimo mayor a Cero", vbInformation, "Aviso"
        txtMontoMin.SetFocus
        Exit Sub
    End If
    
    If txtMontoMax.Text = "0.00" Then
        MsgBox "Digite un Monto Maximo mayor a Cero", vbInformation, "Aviso"
        txtMontoMax.SetFocus
        Exit Sub
    End If
    
    If ValidaRango(Right(CboProd.Text, 3), Right(CboTipoCred.Text, 1), Right(cboMoneda.Text, 1), Trim(txtMontoMin.Text), Trim(txtMontoMax.Text)) = False Then
        MsgBox "El rango especificado ya se encuentra registrado", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    Set D = New COMDCredito.DCOMCredito
    Call D.IngresaRangos(Right(CboProd.Text, 3), Right(CboTipoCred.Text, 1), Right(cboMoneda.Text, 1), Trim(txtMontoMin.Text), Trim(txtMontoMax.Text))
    Set D = Nothing
    
    Call CargaDatosRangos
    HabilitaRangos False
End Sub

Private Function ValidaRango(ByVal psCodProducto As String, ByVal psTipoCred As String, ByVal psMoneda As String, ByVal psMontoMin As String, ByVal psMontoMax As String) As Boolean

ValidaRango = True

If RCredRangos.RecordCount = 0 Then
    ValidaRango = True
    Exit Function
Else
    RCredRangos.MoveFirst
End If

While Not RCredRangos.EOF
    If RCredRangos!cCodProd = psCodProducto And RCredRangos!nTipoCred = CInt(psTipoCred) And RCredRangos!nMoneda = CInt(psMoneda) And RCredRangos!nMontoMin = CDbl(psMontoMin) And RCredRangos!nMontoMax = CDbl(psMontoMax) Then
        ValidaRango = False
        Exit Function
    End If
    RCredRangos.MoveNext
Wend

End Function


Private Sub CmdRangCancel_Click()
    HabilitaRangos False
End Sub

Private Sub CmdRangElim_Click()
Dim D As COMDCredito.DCOMCredito

    If RCredRangos.RecordCount = 0 Then
        MsgBox "No Existen Registros para Realizar esta Operacion", vbInformation, "Aviso"
    End If
    
    If MsgBox("Se van a Eliminar Registros, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set D = New COMDCredito.DCOMCredito
        Call D.EliminarRango(RCredRangos!cCodRango)
        Set D = Nothing
        Call CargaDatosRangos
    End If
    

    HabilitaRangos False

End Sub

Private Sub CmdRangNuevo_Click()
    Call HabilitaRangos(True)
    

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub HabilitaIngresoCargo(ByVal pbHabilita As Boolean)

    If pbHabilita Then 'Habilita Ingreso de Datos
        DGCargos.Height = 1665 '1905
    Else
        DGCargos.Height = 2940
    End If
    DGCargos.Enabled = Not pbHabilita
    txtCargo.Text = ""
    CmdCrgNuevo.Enabled = Not pbHabilita
    CmdCrgEliminar.Enabled = Not pbHabilita
    DGCargos.Enabled = Not pbHabilita
    
End Sub

Private Sub HabilitaResponsableCargo(ByVal pbHabilita As Boolean)

    If pbHabilita Then 'Habilita Ingreso de Datos
        DGUsuCargo.Height = 1335 '1710
    Else
        DGUsuCargo.Height = 2760
    End If
    
    TxtUsu.Text = ""
    LblNomUsu.Caption = ""
    CboAgencia.ListIndex = -1
    CmdUsuNuevo.Enabled = Not pbHabilita
    CmdUsuEliminar.Enabled = Not pbHabilita

End Sub

Private Sub CmdUsuAceptar_Click()
    Dim D  As COMDCredito.DCOMCredito

    If Trim(LblNomUsu.Caption) = "" Then
        MsgBox "Compruebe el Nombre del Usuario Por favor", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If ValidaUsuarioIngresado(RCredCargos!cCodCargo, Trim(TxtUsu.Text)) = False Then
        MsgBox "El Usuario ya esta registrado para el Cargo seleccionado", vbInformation, "Aviso"
        Call CmdUsuCancelar_Click
        Exit Sub
    End If
    
    'ARCV 13-06-2006
    If Format(Right(CboAgencia.Text, 2), "00") <> "00" Then
        If ValidaUsuarioxAgencia(RCredCargos!cCodCargo, Format(Right(CboAgencia.Text, 2), "00")) = False Then
            MsgBox "El Cargo ya esta asignado en la Agencia", vbInformation, "Aviso"
            Call CmdUsuCancelar_Click
            Exit Sub
        End If
    End If
    '-----------------------
    Set D = New COMDCredito.DCOMCredito
    Call D.IngresaCreCargoResponsable(RCredCargos!cCodCargo, UCase(Trim(TxtUsu.Text)), Me.LblNomUsu.Caption, Format(Right(CboAgencia.Text, 2), "00"))
    Set D = Nothing
    
    Call CargaDatosCargosResponsable(RCredCargos!cCodCargo)
    Call HabilitaResponsableCargo(False)

End Sub

Private Function ValidaUsuarioIngresado(ByVal psCodCargo As String, ByVal psCodUsu As String) As Boolean
Dim i As Integer

ValidaUsuarioIngresado = True

If RCredCargosResp.RecordCount = 0 Then
    ValidaUsuarioIngresado = True
    Exit Function
Else
    RCredCargosResp.MoveFirst
End If

While Not RCredCargosResp.EOF
    If RCredCargosResp!cCodCargo = psCodCargo And UCase(RCredCargosResp!cCodUsu) = UCase(psCodUsu) Then
        ValidaUsuarioIngresado = False
        Exit Function
    End If
    RCredCargosResp.MoveNext
Wend

End Function

Private Sub CmdUsuCancelar_Click()
    Call HabilitaResponsableCargo(False)
End Sub

Private Sub CmdUsuCompNom_Click()
Dim U As COMDPersona.UCOMAcceso

    Set U = New COMDPersona.UCOMAcceso
    Me.LblNomUsu.Caption = U.MostarNombre(gsPDC, Trim(TxtUsu.Text))
    Set U = Nothing

    If Trim(LblNomUsu.Caption) = "" Then
        MsgBox "Usuario No Existe", vbInformation, "Aviso"
    End If

End Sub

Private Sub CmdUsuEliminar_Click()
Dim D As COMDCredito.DCOMCredito

    If RCredCargosResp.RecordCount = 0 Then
        MsgBox "No Existen Registros para Realizar esta Operacion", vbInformation, "Aviso"
    End If
    
    If MsgBox("Se van a Eliminar Registros, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set D = New COMDCredito.DCOMCredito
        Call D.EliminarCargoResponsable(RCredCargosResp!cCodCargo, RCredCargosResp!cCodUsu)
        Set D = Nothing
        Call CargaDatosCargosResponsable(RCredCargosResp!cCodCargo)
    End If
    
End Sub

Private Sub CmdUsuNuevo_Click()
    Call HabilitaResponsableCargo(True)
    TxtUsu.SetFocus
    
End Sub

Private Sub DGCargos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call CargaDatosCargosResponsable(RCredCargos!cCodCargo)
End Sub


Private Sub DGRangos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If RCredRangos.RecordCount > 0 And RCredRangos.EOF = False Then
        Call CargaDatosNiveles(RCredRangos!cCodRango)
    End If
End Sub

Private Sub Form_Load()
    
     Call CargaDatosCargos
     Call CargaControles
     Call CargaDatosRangos
     
     CentraForm Me
End Sub


Private Sub TxtCargo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        CmdCrgAceptar.SetFocus
    End If
    
    Call Letras(KeyAscii)
    
End Sub

Private Sub TxtMontoMax_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        CmdRangAceptar.SetFocus
    End If
    
    KeyAscii = NumerosEnteros(KeyAscii)

End Sub

Private Sub TxtMontoMin_GotFocus()
    fEnfoque txtMontoMin
    
End Sub

Private Sub TxtMontoMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMontoMax.SetFocus
    End If
    
    KeyAscii = NumerosEnteros(KeyAscii)
    
End Sub

Private Sub TxtMontoMin_LostFocus()
    If Trim(txtMontoMin.Text) = "" Then
        txtMontoMin.Text = "0.00"
    End If
    
    txtMontoMin.Text = Format(txtMontoMin.Text, "#0.00")
    
End Sub


Private Sub TxtMontoMax_GotFocus()
    fEnfoque txtMontoMax
    
End Sub

Private Sub TxtMontoMax_LostFocus()
    If Trim(txtMontoMax.Text) = "" Then
        txtMontoMax.Text = "0.00"
    End If
    
    txtMontoMax.Text = Format(txtMontoMax.Text, "#0.00")
    
End Sub

Private Sub TxtUsu_Change()
    LblNomUsu.Caption = ""
End Sub

Private Sub TxtUsu_GotFocus()
    fEnfoque TxtUsu
End Sub

Private Sub TxtUsu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdUsuCompNom.SetFocus
    End If
    
    Call Letras(KeyAscii)
    
End Sub

'ARCV 13-06-2006
Private Function ValidaUsuarioxAgencia(ByVal psCodCargo As String, ByVal psCodAge As String) As Boolean
Dim i As Integer

ValidaUsuarioxAgencia = True

If RCredCargosResp.RecordCount = 0 Then
    ValidaUsuarioxAgencia = True
    Exit Function
Else
    RCredCargosResp.MoveFirst
End If

While Not RCredCargosResp.EOF
    If RCredCargosResp!cCodCargo = psCodCargo And RCredCargosResp!cAgeCod = psCodAge Then
        ValidaUsuarioxAgencia = False
        Exit Function
    End If
    RCredCargosResp.MoveNext
Wend

End Function

