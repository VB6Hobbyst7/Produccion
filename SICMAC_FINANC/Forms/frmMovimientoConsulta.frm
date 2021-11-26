VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5F774E03-DB36-4DFC-AAC4-D35DC9379F2F}#1.1#0"; "VertMenu.ocx"
Begin VB.Form frmMovimientoConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos: Consulta"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
   Icon            =   "frmMovimientoConsulta.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "&Detalle ..."
      Height          =   375
      Left            =   8760
      TabIndex        =   40
      Top             =   5940
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdRecibo 
      Caption         =   "&Documento"
      Height          =   375
      Left            =   3885
      TabIndex        =   34
      ToolTipText     =   "Movimientos a los que Referencia este Asiento"
      Top             =   5940
      Width           =   1185
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "&Anterior"
      Height          =   375
      Left            =   5100
      TabIndex        =   33
      ToolTipText     =   "Movimientos a los que Referencia este Asiento"
      Top             =   5940
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   "&Siguiente"
      Height          =   375
      Left            =   6315
      TabIndex        =   32
      ToolTipText     =   "Movimientos que Referencian a este Asiento"
      Top             =   5940
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdBuscar 
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
      Height          =   615
      Left            =   10260
      Picture         =   "frmMovimientoConsulta.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Asiento"
      Height          =   375
      Left            =   7530
      TabIndex        =   15
      Top             =   5940
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   9990
      TabIndex        =   16
      Top             =   5940
      Width           =   1200
   End
   Begin VertMenu.VerticalMenu MnuOpe 
      Height          =   6345
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   11192
      MenuCaption1    =   "Buscar por ..."
      MenuItemsMax1   =   5
      MenuItemIcon11  =   "frmMovimientoConsulta.frx":040C
      MenuItemCaption11=   "Operación"
      MenuItemIcon12  =   "frmMovimientoConsulta.frx":0726
      MenuItemCaption12=   "Agencia"
      MenuItemIcon13  =   "frmMovimientoConsulta.frx":0A40
      MenuItemCaption13=   "Documento"
      MenuItemIcon14  =   "frmMovimientoConsulta.frx":0D5A
      MenuItemCaption14=   "Cuenta Contable"
      MenuItemIcon15  =   "frmMovimientoConsulta.frx":1074
      MenuItemCaption15=   "Nro. Movimiento"
   End
   Begin VB.Frame fraFechaBusca 
      Height          =   705
      Left            =   7140
      TabIndex        =   24
      Top             =   0
      Width           =   3045
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   390
         TabIndex        =   12
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFecha2 
         Height          =   315
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         Caption         =   "Al"
         Height          =   225
         Left            =   1590
         TabIndex        =   26
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Del"
         Height          =   225
         Left            =   90
         TabIndex        =   25
         Top             =   300
         Width           =   405
      End
   End
   Begin TabDlg.SSTab tMov 
      Height          =   5055
      Left            =   1500
      TabIndex        =   35
      Top             =   810
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   8916
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Movimiento"
      TabPicture(0)   =   "frmMovimientoConsulta.frx":160E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label20"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblEstado"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label16"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblPersCod"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblPersNombre"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label15"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblOpeDesc"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblOpecod"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label14"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label18"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblPersRuc"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblMovNroModifica"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fgDoc"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "dgMov"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtMovDesc"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Asiento Contable"
      TabPicture(1)   =   "frmMovimientoConsulta.frx":162A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape1"
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(4)=   "LblTotHS"
      Tab(1).Control(5)=   "Label10"
      Tab(1).Control(6)=   "LblTotDS"
      Tab(1).Control(7)=   "LblTotDD"
      Tab(1).Control(8)=   "LblTotHD"
      Tab(1).Control(9)=   "dgMovAsiento"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Otros datos"
      TabPicture(2)   =   "frmMovimientoConsulta.frx":1646
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox txtMovDesc 
         Height          =   345
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   4410
         Width           =   8865
      End
      Begin MSDataGridLib.DataGrid dgMov 
         Height          =   3765
         Left            =   150
         TabIndex        =   37
         Top             =   540
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   6641
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   6
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "cMovNro"
            Caption         =   "Nro.Movimiento"
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
            DataField       =   "nMovEstado"
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
         BeginProperty Column02 
            DataField       =   "cOpeCod"
            Caption         =   "Operación"
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
            DataField       =   "nMovMonto"
            Caption         =   "Monto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   " #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "cMovDesc"
            Caption         =   "Glosa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "nMovFlag"
            Caption         =   "Flag"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "nMovNro"
            Caption         =   "Nro.Mov"
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
            DataField       =   "MCerrado"
            Caption         =   "M.Cerrado"
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
            MarqueeStyle    =   3
            BeginProperty Column00 
               Alignment       =   2
               DividerStyle    =   6
               ColumnWidth     =   2459.906
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               DividerStyle    =   6
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               DividerStyle    =   6
               Locked          =   -1  'True
               ColumnWidth     =   2775.118
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column07 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgMovAsiento 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   38
         Top             =   420
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   6
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "cMovNro"
            Caption         =   "Nro.Movimiento"
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
            DataField       =   "nMovItem"
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
         BeginProperty Column02 
            DataField       =   "cCtaContCod"
            Caption         =   "Cuenta Contable"
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
            DataField       =   "cCtaContDesc"
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
            DataField       =   "nDebe"
            Caption         =   "Debe MN"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "nHaber"
            Caption         =   "Haber MN"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "nDebeME"
            Caption         =   "Debe ME"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "nHaberME"
            Caption         =   "Haber ME"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   2
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               DividerStyle    =   6
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column02 
               DividerStyle    =   6
               ColumnWidth     =   1635.024
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   2099.906
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               DividerStyle    =   6
               Locked          =   -1  'True
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               DividerStyle    =   6
               Locked          =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
         EndProperty
      End
      Begin Sicmact.FlexEdit fgDoc 
         Height          =   1395
         Left            =   3660
         TabIndex        =   50
         Top             =   2940
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2461
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Tipo-Descripción-Numero-Fecha-nMovNro-cMovNro"
         EncabezadosAnchos=   "350-600-2500-1600-1200-0-0"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3-4-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-0-2-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-L-L-C-L-L"
         FormatosEdit    =   "0-3-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblMovNroModifica 
         Alignment       =   1  'Right Justify
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
         Height          =   225
         Left            =   6090
         TabIndex        =   60
         Top             =   330
         Width           =   3705
      End
      Begin VB.Label lblPersRuc 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   8040
         TabIndex        =   52
         Top             =   1260
         Width           =   1515
      End
      Begin VB.Label Label18 
         Caption         =   "Ruc"
         Height          =   285
         Left            =   7410
         TabIndex        =   53
         Top             =   1350
         Width           =   825
      End
      Begin VB.Label Label14 
         Caption         =   "Operación"
         Height          =   255
         Left            =   3870
         TabIndex        =   59
         Top             =   780
         Width           =   825
      End
      Begin VB.Label lblOpecod 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   4770
         TabIndex        =   58
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblOpeDesc 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   5670
         TabIndex        =   57
         Top             =   720
         Width           =   3885
      End
      Begin VB.Label Label15 
         Caption         =   "Persona"
         Height          =   255
         Left            =   3900
         TabIndex        =   56
         Top             =   1380
         Width           =   825
      End
      Begin VB.Label lblPersNombre 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   4770
         TabIndex        =   55
         Top             =   1650
         Width           =   4785
      End
      Begin VB.Label lblPersCod 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   4770
         TabIndex        =   54
         Top             =   1260
         Width           =   1515
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Height          =   1605
         Left            =   3660
         TabIndex        =   51
         Top             =   540
         Width           =   6135
      End
      Begin VB.Label lblEstado 
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
         Height          =   225
         Left            =   150
         TabIndex        =   9
         Top             =   330
         Width           =   5895
      End
      Begin VB.Label Label20 
         Caption         =   "Documentos"
         Height          =   285
         Left            =   3660
         TabIndex        =   49
         Top             =   2700
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Glosa"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   180
         TabIndex        =   39
         Top             =   4440
         Width           =   525
      End
      Begin VB.Label LblTotHD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -67050
         TabIndex        =   48
         Top             =   3630
         Width           =   1395
      End
      Begin VB.Label LblTotDD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -67065
         TabIndex        =   47
         Top             =   3390
         Width           =   1410
      End
      Begin VB.Label LblTotDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -70200
         TabIndex        =   46
         Top             =   3390
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total M.N. Debe"
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
         Left            =   -71775
         TabIndex        =   45
         Top             =   3390
         Width           =   1425
      End
      Begin VB.Label LblTotHS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -70215
         TabIndex        =   44
         Top             =   3630
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total M.N. Haber"
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
         Left            =   -71760
         TabIndex        =   43
         Top             =   3660
         Width           =   1485
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total M.E. Debe"
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
         Left            =   -68565
         TabIndex        =   42
         Top             =   3390
         Width           =   1410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total M.E. Haber"
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
         Left            =   -68565
         TabIndex        =   41
         Top             =   3660
         Width           =   1470
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000C&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   585
         Left            =   -71880
         Top             =   3330
         Width           =   6465
      End
   End
   Begin VB.Frame fraOpe 
      Height          =   705
      Left            =   1500
      TabIndex        =   17
      Top             =   0
      Width           =   5595
      Begin Sicmact.TxtBuscar txtOpeCod 
         Height          =   330
         Left            =   930
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
         sTitulo         =   ""
      End
      Begin VB.TextBox txtOpeDes 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   3285
      End
      Begin VB.Label lblOpe 
         Caption         =   "Operación"
         Height          =   225
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame fraDoc 
      Height          =   705
      Left            =   1500
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   5595
      Begin VB.TextBox txtDocSerie 
         Height          =   315
         Left            =   1860
         TabIndex        =   3
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox txtDocNro 
         Height          =   315
         Left            =   3540
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Documento :      Serie"
         Height          =   225
         Left            =   180
         TabIndex        =   23
         Top             =   300
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Número"
         Height          =   225
         Left            =   2850
         TabIndex        =   22
         Top             =   300
         Width           =   705
      End
   End
   Begin VB.Frame fraCta 
      Height          =   705
      Left            =   1500
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   5595
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3780
         MaxLength       =   16
         TabIndex        =   6
         Top             =   240
         Width           =   1005
      End
      Begin VB.ComboBox cboFiltro 
         Height          =   315
         ItemData        =   "frmMovimientoConsulta.frx":1662
         Left            =   4800
         List            =   "frmMovimientoConsulta.frx":1678
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   660
      End
      Begin VB.TextBox txtCtaCod 
         Height          =   315
         Left            =   1410
         TabIndex        =   5
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         Height          =   195
         Left            =   3180
         TabIndex        =   31
         Top             =   315
         Width           =   525
      End
      Begin VB.Label Label12 
         Caption         =   "Cuenta Contable"
         Height          =   225
         Left            =   120
         TabIndex        =   30
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.Frame fraMov 
      Height          =   705
      Left            =   1500
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   5595
      Begin VB.TextBox txtMovNro 
         Height          =   315
         Left            =   1860
         TabIndex        =   8
         Top             =   240
         Width           =   3315
      End
      Begin VB.Label Label11 
         Caption         =   "Nro. Movimiento"
         Height          =   225
         Left            =   420
         TabIndex        =   28
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.Frame fraAge 
      Height          =   705
      Left            =   1500
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   5595
      Begin Sicmact.TxtBuscar txtAgeCod 
         Height          =   345
         Left            =   900
         TabIndex        =   10
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
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
         sTitulo         =   ""
      End
      Begin VB.TextBox txtAgeDesc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   11
         Top             =   240
         Width           =   3285
      End
      Begin VB.Label Label3 
         Caption         =   "Agencia"
         Height          =   225
         Left            =   150
         TabIndex        =   20
         Top             =   270
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmMovimientoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbExtorno     As Boolean
Dim lbConsulta    As Boolean
Dim lbEliminaMov  As Boolean
Dim lnTpoBusca    As Integer
Dim lsUltMovNro   As String
Dim rsMov         As ADODB.Recordset

Public Sub Inicio(pbConsulta As Boolean, pbExtorno As Boolean, Optional pbEliminaMov As Boolean, Optional pnTpoBusca As Integer = 1)
lbConsulta = pbConsulta
lbExtorno = pbExtorno
lbEliminaMov = pbEliminaMov
lnTpoBusca = pnTpoBusca
   'lnTpoBusca = 2 Muestra Eliminados
Me.Show 1
End Sub

Private Sub cboFiltro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtFecha.SetFocus
End If
End Sub

Private Sub cmdAnterior_Click()
Dim clsAsiento As New NContAsientos
Dim sBusCond As String
If rsMov.EOF Then
   Exit Sub
End If
sBusCond = " M.nMovNro IN (SELECT nMovNroRef FROM MovRef WHERE nMovNro = " & rsMov!nMovNro & " ) "
Set rsMov = clsAsiento.GetMovimientoConsulta(sBusCond, "", "", "", "", "", "")
Set dgMov.DataSource = rsMov
Set clsAsiento = Nothing
MousePointer = 0
dgMov.SetFocus
End Sub

Private Sub cmdbuscar_Click()
Dim sOpeCond As String
Dim sAgeCond As String
Dim sMovCond As String
Dim sDocCond As String
Dim sCtaCond As String
Dim sFecCond As String
Dim sBusCond As String
sAgeCond = ""
sOpeCond = ""
sMovCond = ""
sDocCond = ""
sCtaCond = ""
sFecCond = ""
MousePointer = 11
   If fraOpe.Visible And txtOpeCod.Text <> "" Then
      sOpeCond = " and ot.cOpeGruCod = '" & txtOpeCod.Text & "' "
   End If
   If fraAge.Visible And txtAgeCod.Text <> "" Then
      sAgeCond = " and cMovNro LIKE '______________" & gsCodCMAC & txtAgeCod.Text & "%' "
   End If
   If fraDoc.Visible And txtDocSerie <> "" Or txtDocNro <> "" Then
      sDocCond = " and md.cDocNro like '" & txtDocSerie & "%" & txtDocNro & "%' "
   End If
   If fraCta.Visible And txtCtaCod <> "" Then
      sCtaCond = " and mc.cCtaContCod LIKE '" & txtCtaCod & "%' " & IIf(nVal(txtImporte) <> 0, " and ABS(nMovImporte) " & cboFiltro & nVal(txtImporte) & " ", "")
   End If
   If fraMov.Visible And txtMovNro <> "" Then
      sMovCond = " cMovNro LIKE '" & txtMovNro & "%'"
   End If
   If fraFechaBusca.Visible Then
      If Trim(txtFecha) <> "/  /" And Trim(txtFecha2) <> "/  /" Then
         sFecCond = " LEFT(M.cMovNro,8) BETWEEN '" & Format(txtFecha, gsFormatoMovFecha) & "' and " _
                  & " '" & Format(txtFecha2, gsFormatoMovFecha) & "' "
      ElseIf Trim(txtFecha) <> "/  /" Then
         sFecCond = " M.cMovNro LIKE '" & Format(txtFecha, gsFormatoMovFecha) & "%' "
      ElseIf Trim(txtFecha2) <> "/  /" Then
         sFecCond = " M.cMovNro LIKE '" & Format(txtFecha2, gsFormatoMovFecha) & "%' "
      Else
         sFecCond = " M.cMovNro LIKE '_%' "
      End If
   End If
Dim clsAsiento As New NContAsientos
Set rsMov = clsAsiento.GetMovimientoConsulta(sBusCond, sOpeCond, sAgeCond, sDocCond, sCtaCond, sMovCond, sFecCond, Format(gsMesCerrado, gsFormatoMovFecha))
Set dgMov.DataSource = rsMov
Set clsAsiento = Nothing
MousePointer = 0
End Sub

Private Sub MuestraAsiento(psMovNro As String)
Dim sMovCond As String
Dim clsAsiento As New NContAsientos
Dim rsAsiento As ADODB.Recordset
sMovCond = ""
MousePointer = 11
sMovCond = " cMovNro = '" & psMovNro & "'"
Set rsAsiento = clsAsiento.GetAsientoConsulta("", "", "", "", "", sMovCond, "", gsMesCerrado)
Set dgMovAsiento.DataSource = rsAsiento
Set clsAsiento = Nothing
MousePointer = 0
End Sub

Private Sub cmdImprimir_Click()
Dim oMov As New NContImprimir
Dim sTexto As String
'sTexto = oMov.ImprimeAsientoContable(lsUltMovNro, gnLinPage, gnColPage)
'EnviaPrevio sTexto, "Asiento Contable", False
ImprimeAsientoContable lsUltMovNro, , , , , , , , , , , , 1
End Sub

Private Sub cmdRecibo_Click()
Dim oConImp As NContImprimir
Dim lsTexto As String
Dim sSql As String
Dim prs  As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Set oConImp = New NContImprimir
Set prs = New ADODB.Recordset
If fgDoc.TextMatrix(fgDoc.Row, 1) = "" Then
    MsgBox "No se registro Documento en Operación", vbInformation, "¡Aviso!"
    Exit Sub
End If
oCon.AbreConexion
If fgDoc.TextMatrix(fgDoc.Row, 1) = TpoDocRecEgreso Then
    sSql = "SELECT mg.nMovNro, mg.cPersCod, p.cPersNombre, md.cDocNro, md.dDocFecha, m.cMovDesc, m.cOpecod , mc.nMovMonto nMovImporte " _
         & "FROM MovGasto mg JOIN Mov m ON m.nMovNro = mg.nMovNro JOIN Persona p ON p.cPersCod = mg.cPersCod " _
         & "     JOIN MovDoc md ON md.nMovNro = mg.nMovNro JOIN MovCont mc ON mc.nMovNro = mg.nMovNro " _
         & "WHERE mg.nMovNro = " & rsMov!nMovNro & " and md.nDocTpo IN (67," & TpoDocRecEgreso & ")" _
         & "  "
    Set prs = oCon.CargaRecordSet(sSql)
    If Not prs.EOF Then
        lsTexto = oConImp.ImprimeReciboEgresos(gnColPage, rsMov!cMovNro, prs!cMovDesc, GetFechaMov(rsMov!cMovNro, True), gsNomCmac, prs!cOpeCod, False, "", ArendirAtencion, "", prs!cDocNro, _
                  prs!dDocFecha, prs!cPersCod, prs!cPersNombre, "", prs!nMovImporte)
    End If
    RSClose prs
End If
If fgDoc.TextMatrix(fgDoc.Row, 1) = TpoDocRecArendirCuenta Then
    Dim oNContImprimir As NContImprimir
    Set oNContImprimir = New NContImprimir
    lsTexto = oNContImprimir.ImprimeReciboARendir(rsMov!cMovNro, gnColPage, gsInstCmac, gsNomCmac, gsNomCmacRUC)
    Set oNContImprimir = Nothing
End If
If lsTexto <> "" Then
    EnviaPrevio lsTexto, Me.Caption, gnLinPage
Else
    MsgBox "Sistema no emite Formato del Tipo de Documento de la Operación", vbInformation, "¡Aviso!"
End If
oCon.CierraConexion
Set oCon = Nothing

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSiguiente_Click()
Dim clsAsiento As New NContAsientos
Dim sBusCond As String
If rsMov.EOF Then
   Exit Sub
End If
sBusCond = " M.nMovNro IN (SELECT nMovNro FROM MovRef WHERE nMovNroRef = " & rsMov!nMovNro & " ) "
Set rsMov = clsAsiento.GetMovimientoConsulta(sBusCond, "", "", "", "", "", "")
Set dgMov.DataSource = rsMov
Set clsAsiento = Nothing
MousePointer = 0
dgMov.SetFocus
End Sub

Private Sub dgMov_HeadClick(ByVal ColIndex As Integer)
If Not rsMov Is Nothing Then
   If Not rsMov.EOF Then
      rsMov.Sort = dgMov.Columns(ColIndex).DataField
   End If
End If
End Sub

Private Sub dgMov_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If rsMov Is Nothing Then
   Exit Sub
End If
RefrescaDatos
End Sub

Private Sub Form_Load()
Dim clsOpe As New DOperacion
txtOpeCod.rs = clsOpe.CargaOpeGru
Set clsOpe = Nothing
cmdAnterior.Visible = True
cmdSiguiente.Visible = True
CentraForm Me
If lbConsulta Then
   Me.Caption = "Asientos Contables: Consulta"
   If lnTpoBusca = 2 Then
      Me.Caption = Me.Caption & " de Movimientos Eliminados"
   End If
End If
If lbExtorno Then
   Me.Caption = "Asientos Contables: Extorno con Generación de Asiento"
End If
If lbEliminaMov Then
   Me.Caption = "Asientos Contables: Extorno con Eliminación de Asiento"
End If
Dim clsRHArea As New DActualizaDatosArea
txtAgeCod.rs = clsRHArea.GetAgencias
Set clsRHArea = Nothing

cboFiltro.ListIndex = 0
lsUltMovNro = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsMov
End Sub

Private Sub MnuOpe_MenuItemClick(MenuNumber As Long, MenuItem As Long)
MnuOpe.MenuItemCur = MenuItem
Select Case MenuItem
   Case 1:  ActivaFrame True, False, False, False, False
            fraFechaBusca.Visible = True
            txtOpeCod.SetFocus
   Case 2:  ActivaFrame False, True, False, False, False
            fraFechaBusca.Visible = True
            txtAgeCod.SetFocus
   Case 3:  ActivaFrame False, False, True, False, False
            fraFechaBusca.Visible = True
            txtDocSerie.SetFocus
   Case 4:  ActivaFrame False, False, False, True, False
            fraFechaBusca.Visible = True
            txtCtaCod.SetFocus
   Case 5:  ActivaFrame False, False, False, False, True
            fraFechaBusca.Visible = False
            txtMovNro.SetFocus
End Select
End Sub

Private Sub txtAgeCod_EmiteDatos()
txtAgeDesc = txtAgeCod.psDescripcion
If txtAgeDesc <> "" And txtFecha.Visible Then
   txtFecha.SetFocus
End If
End Sub

Private Sub txtCtaCod_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtImporte.SetFocus
End If
End Sub

Private Sub txtDocNro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdBuscar.SetFocus
End If
End Sub

Private Sub txtDocSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtDocNro.SetFocus
End If
End Sub

Private Sub txtFecha_GotFocus()
txtFecha.SelStart = 0
txtFecha.SelLength = Len(txtFecha)
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Trim(txtFecha) = "/  /" Then
   Else
      If ValidaFecha(txtFecha) <> "" Then
         MsgBox "Fecha no Valida...!"
         Exit Sub
      End If
   End If
   txtFecha2.SetFocus
End If
End Sub

Private Sub txtFecha2_GotFocus()
txtFecha2.SelStart = 0
txtFecha2.SelLength = Len(txtFecha2)
End Sub

Private Sub txtFecha2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Trim(txtFecha2) = "/  /" Then
   Else
      If ValidaFecha(txtFecha2) <> "" Then
         MsgBox "Fecha no Valida...!"
         Exit Sub
      End If
   End If
   cmdBuscar.SetFocus
End If
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtImporte, KeyAscii, 16, 2)
If KeyAscii = 13 Then
   txtImporte = Format(txtImporte, gsFormatoNumeroView)
   cboFiltro.SetFocus
End If
End Sub

Private Sub txtMovNro_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   cmdBuscar.SetFocus
End If
End Sub

Private Sub txtOpeCod_EmiteDatos()
txtOpeDes = txtOpeCod.psDescripcion
If txtOpeCod <> "" And txtFecha.Visible Then
   txtFecha.SetFocus
End If
End Sub

'**************** FUNCIONES PRIVADA *************
Private Sub ActivaFrame(pbOpe As Boolean, pbAge As Boolean, pbDoc As Boolean, pbCta As Boolean, pbMov As Boolean)
fraOpe.Visible = pbOpe
fraAge.Visible = pbAge
fraDoc.Visible = pbDoc
fraCta.Visible = pbCta
fraMov.Visible = pbMov
End Sub

Private Sub RefrescaDatos()
Dim lsEstado As String
If rsMov Is Nothing Then
   Exit Sub
End If
If Not rsMov.EOF And Not rsMov.BOF Then
   If rsMov!cMovNro <> lsUltMovNro Then
      lblOpecod = rsMov!cOpeCod
      
      MuestraDatosMov rsMov!cMovNro
      Sumas
      lsEstado = "MOVIMIENTO "
      Select Case rsMov!nMovEstado
        Case gMovEstContabNoContable: lsEstado = lsEstado & "NO-CONTABLE "
        Case gMovEstContabPendiente: lsEstado = lsEstado & "PENDIENTE "
        Case gMovEstContabRechazado: lsEstado = lsEstado & "RECHAZADO "
        Case gMovEstLogIngBienAceptado: lsEstado = lsEstado & "ING.BIEN ACEPTADO "
        Case gMovEstLogIngBienRechazado: lsEstado = lsEstado & "ING.BIEN RECHAZADO "
        Case gMovEstLogSaleBienAlmacen: lsEstado = lsEstado & "SALIDA DE ALMACEN "
        Case 14: lsEstado = lsEstado & "NO-CONTABLE "
      End Select
      Select Case rsMov!nMovFlag
         Case gMovFlagEliminado: lsEstado = lsEstado & "ELIMINADO"
         Case gMovFlagDeExtorno: lsEstado = lsEstado & "DE EXTORNO"
         Case gMovFlagExtornado: lsEstado = lsEstado & "EXTORNADO"
         Case gMovFlagModificado: lsEstado = lsEstado & "MODIFICADO"
      End Select
      If Not lsEstado = "MOVIMIENTO " Then
         lblEstado.Caption = lsEstado
      Else
         lblEstado.Caption = ""
      End If
   End If
   MuestraAsiento rsMov!cMovNro
Else
Exit Sub
End If
End Sub

Private Sub MuestraDatosMov(vMovNro As String)
Dim prs As New ADODB.Recordset
Dim oMov As New DMov
txtMovDesc = ""
Set prs = oMov.CargaMovOpeAsiento(0, vMovNro)
If Not prs.EOF Then
   lblOpeDesc = prs!cOpeDesc
   txtMovDesc = prs!cMovDesc
   lsUltMovNro = prs!cMovNro
   lblPersCod = prs!cPersCod
   lblPersRuc = prs!cRuc
   lblPersNombre = prs!cPersNombre
   If Not IsNull(prs!cMovNroModifica) Then
      lblMovNroModifica = prs!cMovNroModifica
   Else
      lblMovNroModifica = ""
   End If
End If
RSClose prs
fgDoc.Rows = 2
fgDoc.Clear
fgDoc.FormaCabecera
fgDoc.rsFlex = oMov.CargaMovDocAsiento(rsMov!nMovNro)
Set oMov = Nothing
End Sub

Private Sub Sumas()
Dim prs As New ADODB.Recordset
Dim oMov As New DMov
LblTotDS.Caption = ""
LblTotDD.Caption = ""
LblTotHS.Caption = ""
LblTotHD.Caption = ""
Set prs = oMov.CargaSumaMovAsiento(lsUltMovNro)
If Not prs.EOF Then
   LblTotDS.Caption = Format(prs!nDebe, gsFormatoNumeroView)
   LblTotDD.Caption = Format(prs!nDebeME, gsFormatoNumeroView)
   LblTotHS.Caption = Format(prs!nHaber, gsFormatoNumeroView)
   LblTotHD.Caption = Format(prs!nHaberME, gsFormatoNumeroView)
End If
RSClose prs
Set oMov = Nothing
End Sub


