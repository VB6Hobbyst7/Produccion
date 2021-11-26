VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMntCtasContab 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  CUENTAS CONTABLES: MANTENIMIENTO"
   ClientHeight    =   6210
   ClientLeft      =   735
   ClientTop       =   2040
   ClientWidth     =   10380
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMntCtasContab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pgbCtaCont 
      Height          =   180
      Left            =   120
      TabIndex        =   40
      Top             =   6000
      Visible         =   0   'False
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   318
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "EXCEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   39
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CheckBox chkAnal 
      Caption         =   "Solo Analiticás"
      Height          =   210
      Left            =   8790
      TabIndex        =   24
      Top             =   3420
      Width           =   1455
   End
   Begin VB.CheckBox chkTodo 
      Caption         =   "Toda moneda"
      Height          =   210
      Left            =   8790
      TabIndex        =   25
      Top             =   3690
      Width           =   1455
   End
   Begin VB.CommandButton cmdNuevaAgencia 
      Caption         =   "Nueva A&gencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3060
      Width           =   1245
   End
   Begin VB.CommandButton cmdEstadCambios 
      Caption         =   "Es&tadisticas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2646
      Width           =   1245
   End
   Begin VB.CommandButton cmdInstancia 
      Caption         =   "&Instancias de Objeto"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   5520
      Width           =   1965
   End
   Begin VB.CommandButton cmdDesasignar 
      Caption         =   "&Desasignar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1380
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAsignar 
      Caption         =   "&Asignar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9060
      TabIndex        =   26
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame frmMoneda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   -60
      Width           =   8550
      Begin VB.OptionButton Moneda 
         Caption         =   "[ 0 ]  Dígito Integrador"
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
         Index           =   0
         Left            =   210
         TabIndex        =   11
         Tag             =   "0"
         Top             =   180
         Value           =   -1  'True
         Width           =   2715
      End
      Begin VB.OptionButton Moneda 
         Caption         =   "[ 6 ] No Monetarias Ajustadas"
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
         Index           =   6
         Left            =   3000
         TabIndex        =   16
         Tag             =   "6"
         Top             =   660
         Width           =   3315
      End
      Begin VB.OptionButton Moneda 
         Caption         =   "[ 4 ] De Capital Reajustables"
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
         Index           =   4
         Left            =   3000
         TabIndex        =   15
         Tag             =   "4"
         Top             =   420
         Width           =   3285
      End
      Begin VB.OptionButton Moneda 
         Caption         =   "[ 3 ] De Actualización Constante"
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
         Index           =   3
         Left            =   3000
         TabIndex        =   14
         Tag             =   "3"
         Top             =   180
         Width           =   3345
      End
      Begin VB.OptionButton Moneda 
         Caption         =   "[ 1 ]  Moneda Naci&onal"
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
         Index           =   1
         Left            =   210
         TabIndex        =   12
         Tag             =   "1"
         Top             =   420
         Width           =   2835
      End
      Begin VB.OptionButton Moneda 
         Caption         =   "[ 2 ]  Moneda E&xtranjera"
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
         Index           =   2
         Left            =   210
         TabIndex        =   13
         Tag             =   "2"
         Top             =   660
         Width           =   2895
      End
   End
   Begin MSDataGridLib.DataGrid grdCtas 
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   990
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   4789
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   2
      RowHeight       =   17
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "cCtaContCod"
         Caption         =   "Código"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         Size            =   2
         BeginProperty Column00 
            DividerStyle    =   6
            ColumnAllowSizing=   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1755.213
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   6
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   6224.882
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2232
      Width           =   1245
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9000
      Picture         =   "frmMntCtasContab.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   90
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   8
      Top             =   1818
      Width           =   1245
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      DisabledPicture =   "frmMntCtasContab.frx":09CC
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   7
      Top             =   1404
      Width           =   1245
   End
   Begin MSDataGridLib.DataGrid grdCtaObjs 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   2566
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
         DataField       =   "nCtaObjOrden"
         Caption         =   "#"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cObjetoCod"
         Caption         =   "Cod.Objeto"
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
         DataField       =   "cObjetoDesc"
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
      BeginProperty Column03 
         DataField       =   "nCtaObjNiv"
         Caption         =   "Niv"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "cCtaObjFiltro"
         Caption         =   "Filtro"
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
         DataField       =   "cCtaObjImpre"
         Caption         =   "Impresión"
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
         ScrollBars      =   2
         BeginProperty Column00 
            Alignment       =   2
            DividerStyle    =   6
            ColumnAllowSizing=   0   'False
            WrapText        =   -1  'True
            ColumnWidth     =   269.858
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   6
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column02 
            DividerStyle    =   6
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   4094.929
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            DividerStyle    =   6
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column04 
            DividerStyle    =   6
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1725.165
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   990
      Width           =   1245
   End
   Begin VB.Frame fraImprime 
      Caption         =   "Impresión"
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
      Height          =   2385
      Left            =   8850
      TabIndex        =   17
      Top             =   930
      Visible         =   0   'False
      Width           =   1365
      Begin VB.OptionButton optImpre 
         Caption         =   "&Sin Agencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   19
         Top             =   525
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   22
         Top             =   1590
         Width           =   1155
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   23
         ToolTipText     =   "Cancelar Cambio de Fecha de Asiento"
         Top             =   1950
         Width           =   1155
      End
      Begin VB.TextBox txtCtaCod 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   21
         Top             =   1140
         Width           =   1215
      End
      Begin VB.OptionButton optImpre 
         Caption         =   "&Grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   20
         Top             =   810
         Width           =   915
      End
      Begin VB.OptionButton optImpre 
         Caption         =   "&Todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   18
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame fraAgencia 
      Caption         =   "Agencia"
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
      Height          =   2985
      Left            =   8790
      TabIndex        =   30
      Top             =   720
      Visible         =   0   'False
      Width           =   1485
      Begin VB.TextBox txtAgeNuevaNom 
         Height          =   345
         Left            =   90
         TabIndex        =   38
         Top             =   1770
         Width           =   1275
      End
      Begin VB.TextBox txtAgeRefNom 
         Height          =   345
         Left            =   90
         TabIndex        =   37
         Top             =   810
         Width           =   1275
      End
      Begin Sicmact.TxtBuscar txtAgeRef 
         Height          =   345
         Left            =   90
         TabIndex        =   35
         Top             =   450
         Width           =   1275
         _ExtentX        =   2249
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
      Begin VB.CommandButton cmdCancelAge 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   32
         ToolTipText     =   "Cancelar Cambio de Fecha de Asiento"
         Top             =   2550
         Width           =   1155
      End
      Begin VB.CommandButton cmdGrabaAge 
         Caption         =   "&Generar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   31
         Top             =   2190
         Width           =   1155
      End
      Begin Sicmact.TxtBuscar txtAgeNueva 
         Height          =   345
         Left            =   90
         TabIndex        =   36
         Top             =   1410
         Width           =   1305
         _ExtentX        =   2302
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
      Begin VB.Label Label3 
         Caption         =   "Nueva"
         ForeColor       =   &H000000FF&
         Height          =   165
         Left            =   180
         TabIndex        =   34
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Referencia"
         ForeColor       =   &H000000FF&
         Height          =   165
         Left            =   180
         TabIndex        =   33
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Objetos "
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
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   735
   End
End
Attribute VB_Name = "frmMntCtasContab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSqlCta As String
Dim rsCtaObj As ADODB.Recordset
Dim rsCta As ADODB.Recordset

Dim clsCtaCont As DCtaCont
Dim nOrdenCta As Integer 'Para establecer el Orden de la tablas
Dim lPressMouse As Boolean

Dim WithEvents oImp As NContImprimir
Attribute oImp.VB_VarHelpID = -1
Dim oBarra As clsProgressBar
Dim E As Boolean
Dim lnIndex As Integer
'ARLO20170208****
Dim objPista As COMManejador.Pista
'************


Private Sub ManejaBotonVisible(plOpcion As Boolean, plActivaImp As Boolean, plActivaAge As Boolean)
cmdBuscar.Visible = plOpcion
cmdNuevo.Visible = plOpcion
cmdModificar.Visible = plOpcion
cmdEliminar.Visible = plOpcion
cmdImprimir.Visible = plOpcion

cmdEstadCambios.Visible = plOpcion

cmdNuevaAgencia.Visible = plOpcion
fraImprime.Visible = plActivaImp
chkTodo.Visible = plActivaImp
chkAnal.Visible = plActivaImp

fraAgencia.Visible = plActivaAge
End Sub

Private Sub ManejaBoton(plOpcion As Boolean)
cmdBuscar.Enabled = plOpcion
If Moneda(0).value = True Then
   If (Moneda(0).value Or Moneda(6).value) Or Not plOpcion Then
      cmdNuevo.Enabled = plOpcion
      cmdModificar.Enabled = plOpcion
   End If
End If
cmdEliminar.Enabled = plOpcion
cmdImprimir.Enabled = plOpcion
cmdNuevaAgencia.Enabled = plOpcion
'cmdEstadCambios.Enabled = plOpcion

frmMoneda.Enabled = plOpcion
Moneda(0).Enabled = plOpcion
Moneda(1).Enabled = plOpcion
Moneda(2).Enabled = plOpcion
Moneda(3).Enabled = plOpcion
Moneda(4).Enabled = plOpcion
Moneda(6).Enabled = plOpcion

grdCtas.Enabled = plOpcion
grdCtaObjs.Enabled = plOpcion
cmdAsignar.Enabled = plOpcion
cmdDesasignar.Enabled = plOpcion
chkTodo.Enabled = Not plOpcion
chkAnal.Enabled = Not plOpcion
End Sub

Private Sub RefrescaGrid(npMoneda As Integer)
Set rsCta = clsCtaCont.CargaCtaCont("SubString(cCtaContCod,3,1)='" & Trim(Str(npMoneda)) & "' or Len(rtrim(cCtaContCod))<3", "CtaCont", adLockOptimistic)
Set grdCtas.DataSource = rsCta
End Sub

Private Sub cmdAsignar_Click()
Dim s As String
Dim nOrden As Integer
' SOLO SE ASIGNAN OBJETOS A ULTIMA INSTANCIA DE CUENTA
If Not clsCtaCont.CtaInstancia(Trim(rsCta!cCtaContCod), "CtaCont") Then
   If MsgBox("Cuenta no es Instancia. ¿ Desea continuar ? ", vbQuestion + vbYesNo, "Advertencia") = vbNo Then
      grdCtas.SetFocus
      Exit Sub
   End If
End If
nOrden = 0
If Not rsCtaObj Is Nothing Then
   If Not rsCtaObj.EOF Then
      rsCtaObj.MoveLast
      nOrden = rsCtaObj!nCtaObjOrden
   End If
End If
frmAsignaObj.Inicio rsCta!cCtaContCod, rsCta!cCtaContDesc, "0", nOrden + 1
CargaCtaObjs
grdCtas.SetFocus
End Sub

Private Sub cmdbuscar_Click()
Dim clsBuscar As New ClassDescObjeto
On Error GoTo ErrMsg
ManejaBoton False
clsBuscar.BuscarDato rsCta, nOrdenCta, "Cuenta Contable"
If clsBuscar.lbOk Then
    CargaCtaObjs
End If
nOrdenCta = clsBuscar.gnOrdenBusca
Set clsBuscar = Nothing
ManejaBoton True
grdCtas.SetFocus
Exit Sub
ErrMsg:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdCancelAge_Click()
ManejaBotonVisible True, False, False
ManejaBoton True
End Sub

Private Sub cmdCancelar_Click()
ManejaBotonVisible True, False, False
ManejaBoton True
End Sub

Private Sub cmdDesasignar_Click()
Dim dSQL As String
If rsCtaObj Is Nothing Then
   MsgBox "No se asignaron Objetos a Cuenta...!", vbExclamation, "Aviso"
   grdCtas.SetFocus
   Exit Sub
End If
If RSVacio(rsCtaObj) Then
   MsgBox "No se asignaron Objetos a Cuenta...!", vbExclamation, "Aviso"
   grdCtas.SetFocus
   Exit Sub
End If

If grdCtaObjs.row + 1 = rsCtaObj.RecordCount Then
   If MsgBox(" ¿ Seguro Desasignar Objeto de Cuenta Contable ? ", vbOKCancel, "Confirmación de Eliminación") = vbOk Then
     clsCtaCont.EliminaCtaObj rsCta!cCtaContCod, rsCtaObj!cObjetoCod
     CargaCtaObjs
   End If
Else
   MsgBox "Existen objetos en nivel inferior", vbExclamation, "Error de Desasignación"
End If
If grdCtaObjs.Enabled Then
   grdCtaObjs.SetFocus
Else
   grdCtas.SetFocus
End If
End Sub

Private Sub cmdEliminar_Click()
Dim dSQL As String
Dim Pos As Variant
On Error GoTo DelError
If Not rsCta.BOF And Not rsCta.EOF Then
   If Not clsCtaCont.CtaInstancia(rsCta!cCtaContCod, "CtaCont") Then
      MsgBox " Existen Cuentas en nivel Inferior ...!", vbExclamation, "Aviso de Eliminación"
      grdCtas.SetFocus
      Exit Sub
   End If
   If rsCtaObj.RecordCount > 0 Then
      MsgBox " Existen Objetos asignados a Cuenta ...!", vbExclamation, "Aviso de Eliminación"
      grdCtas.SetFocus
      Exit Sub
   End If
   If MsgBox(" ¿ Seguro de Eliminar Cuenta ? ", vbOKCancel, "Mensaje de Confirmación") = vbOk Then
      
      clsCtaCont.InsertaCtaContHisto rsCta!cCtaContCod, rsCta!cCtaContDesc, gsMovNro
      clsCtaCont.EliminaCtaCont Mid(rsCta!cCtaContCod, 1, 2) & IIf(Len(rsCta!cCtaContCod) > 2, "_", "") & Mid(rsCta!cCtaContCod, 4, 20), "CtaCont"
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantCctaCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Se Elimino la Cuenta Contable : " & rsCta!cCtaContCod
            Set objPista = Nothing
            '*******
      rsCta.Delete adAffectCurrent
      
   End If
   grdCtas.SetFocus
Else
   MsgBox "No existen Cuentas para eliminar...", vbInformation, "Error de Eliminación"
End If
Exit Sub
DelError:
 MsgBox TextErr(Err.Description), vbExclamation, "Error de Eliminación"
End Sub

Private Sub cmdEstadCambios_Click()
frmMntCtasContabCambios.Show 1, Me
End Sub

Private Sub cmdExcel_Click()
E = True
    Dim liLineas As Integer
    Dim I As Integer
    
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet

    Dim lbConexion As Boolean
    Dim lbExisteHoja  As Boolean
    Dim lsNomHoja As String
    Dim glsArchivo As String
      
 
    glsArchivo = "CuentasContables" & Format(Now, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    Set xlHoja1 = xlLibro.Worksheets.Add

    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    xlHoja1.PageSetup.Orientation = xlLandscape
    
    Call imprimirCtaCont(xlLibro, xlHoja1, xlAplicacion)
            
    xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        'Cierra el libro de trabajo
    xlLibro.Close
        ' Cierra Microsoft Excel con el método Quit.
    xlAplicacion.Quit
    
    MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
    CargaArchivo App.path & "\SPOOLER\" & glsArchivo, App.path & "\SPOOLER\"
    
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantCctaCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Imprimio en Excel las Cuentas Contables "
            Set objPista = Nothing
            '*******

'imprimirCtaCont
End Sub
Public Sub imprimirCtaCont(xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, xlAplicacion As Excel.Application)
Dim I As Integer
Dim HojasExcel      As Integer 'numero de hojas de Excel a usar para mostrar las Cuentas contables
Dim RSTEMP As ADODB.Recordset
Dim nHojasExcel As Integer

  Dim cMES As String
    Dim cnomhoja As String
    Dim liLineas As Long
    Dim liLineas1 As Long
    Dim nReg As Integer
    Dim lnTipoCambio As Currency
    Dim glsArchivo As String

Set RSTEMP = New ADODB.Recordset
    
    Set RSTEMP = clsCtaCont.CargaCtaCont("SubString(cCtaContCod,3,1)='" & Trim(Str(lnIndex)) & "' or Len(rtrim(cCtaContCod))<3", "CtaCont", adLockOptimistic)
       
    If RSTEMP Is Nothing Then
        MsgBox "No exite informacion para imprimir", vbInformation, "Aviso"
        Exit Sub
    End If
    Me.pgbCtaCont.Visible = True
    Me.pgbCtaCont.Min = 0
    Me.pgbCtaCont.Max = RSTEMP.RecordCount
    
    nHojasExcel = Round(RSTEMP.RecordCount / 65000)
    
For I = 1 To nHojasExcel + 1
        liLineas = 1
        cnomhoja = "CuentasContables" & I
    
    Call ExcelAddHoja(cnomhoja, xlLibro, xlHoja1)
    
            xlAplicacion.Range("A1:A1").ColumnWidth = 2
            xlAplicacion.Range("B1:B1").ColumnWidth = 20
            xlAplicacion.Range("c1:c1").ColumnWidth = 70
          
            xlAplicacion.Range("A1:Z100").Font.Size = 9

            xlHoja1.Cells(1, 1) = gsNomCmac
            xlHoja1.Cells(2, 2) = "L I S T A D O   D E   C U E N T A S   C O N T A B L E S  "
            xlHoja1.Cells(3, 2) = "INFORMACION  AL  " & Format(gdFecSis, "dd/mm/yyyy")

            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 3)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 3)).Merge True
            xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 3)).Merge True
            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 3)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 3)).HorizontalAlignment = xlCenter

            liLineas = 6

            xlHoja1.Cells(liLineas, 2) = "Codigo"
            xlHoja1.Cells(liLineas, 3) = "Descripcion"
         

            xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 3)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 3)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 3)).Borders.LineStyle = 1
            xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 3)).Interior.Color = RGB(159, 206, 238)
            liLineas = liLineas + 1
    
        Do Until RSTEMP.EOF   '.Bookmark = 65000 * i
        'Do While Not rsTemp.EOF
             
                xlHoja1.Cells(liLineas, 2) = RSTEMP(0)
                xlHoja1.Cells(liLineas, 3) = RSTEMP(1)
                liLineas = liLineas + 1
                Me.pgbCtaCont.value = RSTEMP.Bookmark
                RSTEMP.MoveNext
                
                'If rsTemp.Bookmark = 65000 * i Or rsTemp.EOF Then
                If liLineas = 65000 Then
                   Exit Do
                End If
                      
         Loop
Next
    Me.pgbCtaCont.Visible = False
End Sub

Private Sub cmdGrabaAge_Click()
Dim sSql As String
If Me.txtAgeRefNom = "" Then
    MsgBox "Nombre de Agencia de Referencia no indicado", vbInformation, "¡Aviso!"
    Me.txtAgeRef.SetFocus
    Exit Sub
End If

If Me.txtAgeNuevaNom = "" Then
    MsgBox "Nombre de Nueva Agencia no indicado", vbInformation, "¡Aviso!"
    Me.txtAgeNueva.SetFocus
    Exit Sub
End If

Dim oCon As New DConecta
Dim oMov As New DMov
gsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

sSql = "INSERT CtaCont(cCtaContCod, cctacontdesc, cUltimaActualizacion) " _
     & "SELECT distinct LEFT( c.cCtaContCod,len(c.cctacontcod)-2) + '" & txtAgeNueva & "', '" & Me.txtAgeNuevaNom & "', '" & gsMovNro & "' " _
     & "FROM ctacont c left join ctacont c1 on c1.cctacontcod = lefT(c.cCtaContCod,len(c.cctacontcod)-2) + '" & txtAgeNueva & "' " _
     & "WHERE c.cctacontdesc like '" & Trim(Me.txtAgeRefNom) & "' and c1.cCtaContCod is NULL "
oCon.AbreConexion
oCon.Ejecutar sSql
oCon.CierraConexion
Set oCon = Nothing
MsgBox "Cuentas Contables generadas Satisfactoriamente", vbInformation, "¡Aviso!"
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantCctaCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Cuentas Contables generadas Satisfactoriamente |Agencia Referencia : " & txtAgeRefNom.Text & "|Agencia Nueva : " & txtAgeNuevaNom.Text
            Set objPista = Nothing
            '*******
ManejaBotonVisible True, False, False
ManejaBoton True
Me.Enabled = True
End Sub

Private Sub cmdImprimir_Click()
ManejaBoton False
ManejaBotonVisible False, True, False
If optImpre(0).value Then
   optImpre(0).SetFocus
ElseIf optImpre(1).value Then
   optImpre(1).SetFocus
Else
   optImpre(2).SetFocus
End If
End Sub

Private Sub cmdAplicar_Click()
Dim sSql As String
Dim rs   As New ADODB.Recordset
Dim N    As Integer
Dim sTexto As String, lsImpre As String
Dim sNomFile As String
MousePointer = 11
Me.Enabled = False
For N = 0 To Moneda.Count
   If Moneda(IIf(N = 5, 6, N)).value Then
        If optImpre(0).value Then
            If chkTodo.value = vbChecked Then
               sSql = "( cCtaContCod like '__[^0]%' "
            Else
               sSql = "( SubString(cCtaContCod,3,1)='" & Moneda(IIf(N = 5, 6, N)).Tag & "' "
            End If
            sSql = sSql & " or Len(rtrim(cCtaContCod))<3 ) "
           sNomFile = "PLAN_Todo.txt"
        ElseIf optImpre(2).value Then
            If chkTodo.value = vbChecked Then
               sSql = "cCtaContCod like '" & Left(txtCtaCod, 2) & "[^0]" & Mid(txtCtaCod, 4, 20) & "%' "
            Else
               sSql = "cCtaContCod LIKE '" & txtCtaCod & "%' and SubString(cCtaContCod,3,1) = '" & Moneda(IIf(N = 5, 6, N)).Tag & "' "
            End If
           sNomFile = "PLAN_Grupo.txt"
        Else
            If chkTodo.value = vbChecked Then
               sSql = ""
            Else
               sSql = "cCtaContCod LIKE '__" & Moneda(IIf(N = 5, 6, N)).Tag & "%' and "
            End If
            sSql = sSql & " NOT CCTACONTDESC LIKE 'AGENCIA %' and NOT cCtaContDesc LIKE 'OFICINA ESPECIAL%' " _
                        & " And NOT cCtaContDesc LIKE 'OFICINA ESP.%' and NOT cCtaContDesc LIKE '%OFIC. ESPEC.%' " _
                        & " And NOT cCtaContDesc LIKE 'SEDE INSTITUCIONAL%' and NOT cCtaContDesc LIKE 'OFIC. ESPECIAL%' " _
                        & " And NOT cCtaContDesc LIKE 'OFIC.ESPECIAL%' and NOT cCtaContDesc LIKE 'OFICINA%'" _
                        & " And NOT cCtaContDesc LIKE 'CHICLAYO%' and NOT cCtaContDesc LIKE 'AG. CAJAMARCA%' "
           sNomFile = "PLAN_SAgencia.txt"
        End If
        If chkAnal.value = vbChecked Then
           sSql = sSql & " and len(cctacontcod) = (select Max(Len(cCtaContCod)) from CtaCont c WHERE c.cCtaContCod LIKE ctacont.cCtaContCod + '%' ) "
        End If
        sSql = sSql & " Order By cCtaContCod "
      
      Exit For
   End If
Next
Set rs = clsCtaCont.CargaCtaCont(sSql, "CtaCont")
If rs.EOF Then
   MsgBox "No Existen Cuentas Contables Registradas", vbInformation, "Aviso"
   Me.Enabled = True
   RSClose rs
   MousePointer = 0
   Exit Sub
End If
MousePointer = 0
   Set oImp = New NContImprimir
   oImp.Inicio gsNomCmac, gsCodAge, Format(gdFecSis, gsFormatoFechaView)
   lsImpre = oImp.ImprimePlanContable(rs, "CtaCont", gnLinPage)
   Set oImp = Nothing
   RSClose rs
MousePointer = 0
EnviaPrevio lsImpre, "Cuenta Contables : Reporte ", gnLinPage, False
ManejaBotonVisible True, False, False
ManejaBoton True
Me.Enabled = True
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantCctaCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Imprimio las Cuentas Contables "
            Set objPista = Nothing
            '*******
End Sub

Private Sub cmdInstancia_Click()
If Not rsCtaObj Is Nothing Then
If Not rsCtaObj.EOF Then
    frmMntCtasContInstancia.Inicia rsCta!cCtaContCod, rsCta!cCtaContDesc, rsCtaObj!cObjetoCod, rsCtaObj!cObjetoDesc, rsCtaObj!nObjetoNiv, rsCtaObj!nCtaObjNiv, rsCtaObj!cCtaObjFiltro, rsCtaObj!nCtaObjOrden
    CargaCtaObjs
End If
End If
End Sub

Private Sub cmdModificar_Click()
'Dim lsCtaCod As String
'ManejaBoton False
'lsCtaCod = rsCta!cCtaContCod
'If gsProyectoActual = "H" Then
'    frmMntCtasContNuevoHC.Inicia False, rsCta!cCtaContCod, rsCta!cCtaContDesc, , Moneda(0).value
'Else
'    frmMntCtasContNuevo.Inicia False, rsCta!cCtaContCod, rsCta!cCtaContDesc, , Moneda(0).value
'End If
'If frmMntCtasContNuevo.OK Then
'   RefrescaGrid IIf(Moneda(0).value, 0, 6)
'   rsCta.Find "cCtaContCod = '" & lsCtaCod & "'"
'End If
'ManejaBoton True
Dim lsCtaCod As String
ManejaBoton False
lsCtaCod = rsCta!cCtaContCod
'If gsProyectoActual = "H" Then
    frmMntCtasContNuevoHC.Inicia False, rsCta!cCtaContCod, rsCta!cCtaContDesc, , Moneda(0).value
    If frmMntCtasContNuevoHC.OK Then
        RefrescaGrid IIf(Moneda(0).value, 0, 6)
        rsCta.Find "cCtaContCod = '" & lsCtaCod & "'"
    End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantCctaCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "2", "Se Modifico la Cuenta Contable : " & frmMntCtasContNuevoHC.cCtaContCod
            Set objPista = Nothing
            '*******
'Else
'    frmMntCtasContNuevo.Inicia False, rsCta!cCtaContCod, rsCta!cCtaContDesc, , Moneda(0).value
'    If frmMntCtasContNuevo.OK Then
'        RefrescaGrid IIf(Moneda(0).value, 0, 6)
'        rsCta.Find "cCtaContCod = '" & lsCtaCod & "'"
'    End If
'End If
ManejaBoton True

End Sub

Private Sub cmdNuevaAgencia_Click()
ManejaBoton False
ManejaBotonVisible False, False, True
txtAgeRef.SetFocus
End Sub

Private Sub cmdNuevo_Click()
Dim Pos As Variant
ManejaBoton False

'If gsProyectoActual = "H" Then
    frmMntCtasContNuevoHC.Inicia True, "", "", , Moneda(0).value
    If frmMntCtasContNuevoHC.OK Then
       RefrescaGrid IIf(Moneda(0).value, 0, 6)
       rsCta.Find "cCtaContCod = '" & frmMntCtasContNuevoHC.cCtaContCod & "'", 0, adSearchForward, 1
    End If
'Else
'    frmMntCtasContNuevo.Inicia True, "", "", , Moneda(0).value
'    If frmMntCtasContNuevo.OK Then
'       RefrescaGrid IIf(Moneda(0).value, 0, 6)
'       rsCta.Find "cCtaContCod = '" & frmMntCtasContNuevo.cCtaContCod & "'", 0, adSearchForward, 1
'    End If
'End If
ManejaBoton True
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Form_Activate()
ManejaBoton True
'grdCtas.SetFocus
End Sub

Private Sub Form_Load()
CentraForm Me
frmMdiMain.Enabled = False
lPressMouse = False
nOrdenCta = 0  'Inicialmente las cuentas se Ordenan por Codigo

Set clsCtaCont = New DCtaCont
RefrescaGrid 0
Set grdCtas.DataSource = rsCta

Dim oAge As New DActualizaDatosArea
txtAgeNueva.rs = oAge.GetAgencias()
txtAgeRef.rs = oAge.GetAgencias()
Set oAge = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set clsCtaCont = Nothing
Set rsCtaObj = Nothing
frmMdiMain.Enabled = True
End Sub

Private Sub grdCtaObjs_GotFocus()
grdCtaObjs.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdCtaObjs_LostFocus()
grdCtaObjs.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub grdCtas_GotFocus()
grdCtas.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdCtas_HeadClick(ByVal ColIndex As Integer)
If Not rsCta Is Nothing Then
   If Not rsCta.EOF Then
      rsCta.Sort = grdCtas.Columns(ColIndex).DataField
   End If
End If
End Sub

Private Sub grdCtas_LostFocus()
grdCtas.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub grdCtas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If E = False Then
If Not LastRow = "" Then
   If Not rsCta.EOF Then
      If LastRow <> rsCta.Bookmark Then
         CargaCtaObjs
      End If
   End If
End If
End If
End Sub

Private Sub Moneda_Click(Index As Integer)
RefrescaGrid Index
lnIndex = Index
CargaCtaObjs
If Index = 0 Or Index = 6 Then
   cmdNuevo.Enabled = True
   cmdModificar.Enabled = True
Else
   cmdNuevo.Enabled = False
   cmdModificar.Enabled = False
End If
End Sub

Private Sub CargaCtaObjs()
If Not rsCta.EOF And Not rsCta.BOF Then
   Set rsCtaObj = clsCtaCont.CargaCtaObj(rsCta!cCtaContCod, , True)
   If rsCtaObj.EOF Then
      cmdInstancia.Enabled = False
   Else
      cmdInstancia.Enabled = True
   End If
   grdCtaObjs.Enabled = True
   cmdDesasignar.Enabled = True
   cmdAsignar.Enabled = True
Else
   Set rsCtaObj = Nothing
   cmdInstancia.Enabled = False
   grdCtaObjs.Enabled = False
   cmdDesasignar.Enabled = False
   cmdAsignar.Enabled = False
End If
Set grdCtaObjs.DataSource = rsCtaObj
End Sub

Private Sub optImpre_Click(Index As Integer)
If Index = 2 Then
   txtCtaCod.Enabled = True
   txtCtaCod.SetFocus
Else
   txtCtaCod.Enabled = False
End If
End Sub

Private Sub txtAgeNueva_EmiteDatos()
Me.txtAgeNuevaNom = txtAgeNueva.psDescripcion
If txtAgeNuevaNom <> "" Then
    Me.cmdGrabaAge.SetFocus
End If
End Sub

Private Sub txtAgeRef_EmiteDatos()
Me.txtAgeRefNom = txtAgeRef.psDescripcion
If txtAgeRefNom <> "" Then
    Me.txtAgeRef.SetFocus
End If
End Sub

Private Sub txtCtaCod_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   cmdAplicar.SetFocus
End If
End Sub

Private Sub oImp_BarraClose()
oBarra.CloseForm Me
Set oBarra = Nothing
End Sub

Private Sub oImp_BarraProgress(value As Variant, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
oBarra.Progress value, psTitulo, psSubTitulo, psTituloBarra, ColorLetras
End Sub

Private Sub oImp_BarraShow(pnMax As Variant)
Set oBarra = New clsProgressBar
oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = pnMax
End Sub



