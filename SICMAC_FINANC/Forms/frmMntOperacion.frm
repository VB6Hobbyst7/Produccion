VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMntOperacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operaciones: Mantenimiento"
   ClientHeight    =   6300
   ClientLeft      =   1935
   ClientTop       =   1710
   ClientWidth     =   9645
   Icon            =   "frmMntOperacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3435
      Left            =   120
      TabIndex        =   18
      Top             =   30
      Width           =   8115
      Begin VB.TextBox txtOpeGruDes 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1170
         TabIndex        =   24
         Top             =   3000
         Width           =   3645
      End
      Begin VB.TextBox txtOpeGruCod 
         Enabled         =   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   23
         Top             =   3000
         Width           =   405
      End
      Begin VB.CheckBox chkVisible 
         Caption         =   "Visible"
         Height          =   225
         Left            =   6960
         TabIndex        =   21
         Top             =   3060
         Width           =   825
      End
      Begin VB.TextBox txtOpeNivel 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   5820
         TabIndex        =   20
         Top             =   3000
         Width           =   495
      End
      Begin MSComctlLib.TreeView tvOpe 
         Height          =   2775
         Left            =   60
         TabIndex        =   0
         Top             =   150
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   4895
         _Version        =   393217
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imglstFiguras"
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Grupo"
         Height          =   285
         Left            =   150
         TabIndex        =   22
         Top             =   3030
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Nivel"
         Height          =   285
         Left            =   5310
         TabIndex        =   19
         Top             =   3060
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8340
      TabIndex        =   6
      ToolTipText     =   "Salir de mantenimiento de operaciones"
      Top             =   2700
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   8340
      TabIndex        =   5
      Top             =   1830
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   8340
      TabIndex        =   4
      ToolTipText     =   "Elimina una operación"
      Top             =   1410
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   8340
      TabIndex        =   3
      ToolTipText     =   "Modifica la descripción de una operación"
      Top             =   990
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   8340
      TabIndex        =   2
      ToolTipText     =   "Agrega una nueva operación"
      Top             =   570
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   8340
      TabIndex        =   1
      ToolTipText     =   "Busca operación por su código"
      Top             =   150
      Width           =   1215
   End
   Begin TabDlg.SSTab pageOpe 
      Height          =   2685
      Left            =   120
      TabIndex        =   7
      Top             =   3570
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   4736
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   670
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Cuentas Contables"
      TabPicture(0)   =   "frmMntOperacion.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grdCta"
      Tab(0).Control(1)=   "cmdQuitCta"
      Tab(0).Control(2)=   "cmdAsigCta"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Objetos"
      TabPicture(1)   =   "frmMntOperacion.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAsigObj"
      Tab(1).Control(1)=   "cmdQuitObj"
      Tab(1).Control(2)=   "grdObj"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "&Documentos"
      TabPicture(2)   =   "frmMntOperacion.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdDoc"
      Tab(2).Control(1)=   "cmdQuitDoc"
      Tab(2).Control(2)=   "cmdAsigDoc"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Bienes"
      TabPicture(3)   =   "frmMntOperacion.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "dgCtaObj"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdAsigBien"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdQuitBien"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin VB.CommandButton cmdQuitBien 
         Caption         =   "&Quitar Cuenta"
         Height          =   345
         Left            =   1980
         TabIndex        =   27
         Top             =   2250
         Width           =   1695
      End
      Begin VB.CommandButton cmdAsigBien 
         Caption         =   "&Asignar Cuenta"
         Height          =   345
         Left            =   150
         TabIndex        =   26
         Top             =   2250
         Width           =   1695
      End
      Begin VB.CommandButton cmdAsigCta 
         Caption         =   "&Asignar Cuenta"
         Height          =   345
         Left            =   -74850
         TabIndex        =   13
         Top             =   2250
         Width           =   1695
      End
      Begin VB.CommandButton cmdQuitCta 
         Caption         =   "&Quitar Cuenta"
         Height          =   345
         Left            =   -73020
         TabIndex        =   12
         Top             =   2250
         Width           =   1695
      End
      Begin VB.CommandButton cmdAsigObj 
         Caption         =   "&Asignar Objeto"
         Height          =   345
         Left            =   -74850
         TabIndex        =   11
         Top             =   2250
         Width           =   1695
      End
      Begin VB.CommandButton cmdQuitObj 
         Caption         =   "&Quitar Objeto"
         Height          =   345
         Left            =   -73020
         TabIndex        =   10
         Top             =   2250
         Width           =   1695
      End
      Begin VB.CommandButton cmdAsigDoc 
         Caption         =   "&Agregar Documento"
         Height          =   375
         Left            =   -74850
         TabIndex        =   9
         Top             =   2250
         Width           =   1695
      End
      Begin VB.CommandButton cmdQuitDoc 
         Caption         =   "&Quitar Documento"
         Height          =   375
         Left            =   -73020
         TabIndex        =   8
         Top             =   2250
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid grdCta 
         Height          =   1665
         Left            =   -74850
         TabIndex        =   14
         Top             =   540
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   2937
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   17
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "cOpeCtaOrden"
            Caption         =   "Orden"
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
            DataField       =   "cCtaContcod"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "Clase"
            Caption         =   "Clase"
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
            DataField       =   "cCtaContN"
            Caption         =   "Cta Exportar"
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
            Size            =   2
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4380.095
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1409.953
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grdObj 
         Height          =   1665
         Left            =   -74850
         TabIndex        =   15
         Top             =   540
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   2937
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
         WrapCellPointer =   -1  'True
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "cOpeObjOrden"
            Caption         =   " #"
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
            DataField       =   "cObjetoCod"
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
            DataField       =   "nOpeObjNiv"
            Caption         =   "Nivel"
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
            DataField       =   "cOpeObjFiltro"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            ScrollBars      =   2
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1530.142
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4215.118
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   1649.764
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid grdDoc 
         Height          =   1665
         Left            =   -74850
         TabIndex        =   16
         Top             =   540
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   2937
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   18
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
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "cDocDesc"
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
         BeginProperty Column01 
            DataField       =   "Estado"
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
            DataField       =   "Metodo"
            Caption         =   "Método"
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
               ColumnWidth     =   3780.284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2819.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1860.095
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgCtaObj 
         Height          =   1695
         Left            =   150
         TabIndex        =   25
         Top             =   510
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   2990
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   21
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "cCtaContCod"
            Caption         =   "Cuenta"
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
         BeginProperty Column02 
            DataField       =   "cBSCod"
            Caption         =   "Bien/Servicio"
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
            DataField       =   "cBSDescripcion"
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
            ScrollBars      =   2
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1725.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1620.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2805.166
            EndProperty
         EndProperty
      End
   End
   Begin RichTextLib.RichTextBox rtfImp 
      Height          =   375
      Left            =   270
      TabIndex        =   17
      Top             =   180
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmMntOperacion.frx":037A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstFiguras 
      Left            =   30
      Top             =   5700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMntOperacion.frx":03FA
            Key             =   "Padre"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMntOperacion.frx":074C
            Key             =   "Hijo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMntOperacion.frx":0A9E
            Key             =   "Hijito"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMntOperacion.frx":0DF0
            Key             =   "Bebe"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMntOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim rd As New ADODB.Recordset
Dim rc As New ADODB.Recordset
Dim ro As New ADODB.Recordset
Dim rb As New ADODB.Recordset

Dim sTexto As String
Dim sOpeDesc As String

Dim lConsulta  As Boolean
Dim lEnfocaChk As Boolean
Dim clsOpe     As DOperacion
Dim WithEvents oImp As NContImprimir
Attribute oImp.VB_VarHelpID = -1
Dim oBarra As clsProgressBar
'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(plConsulta As Boolean)
lConsulta = plConsulta
Me.Show 0, frmMdiMain
End Sub

Private Sub chkVisible_Click()
If lEnfocaChk Then
   If chkVisible.value = vbChecked Then
      chkVisible.value = vbUnchecked
   Else
      chkVisible.value = vbChecked
   End If
End If
End Sub

Private Sub chkVisible_GotFocus()
tvOpe.SetFocus
lEnfocaChk = True
End Sub

Private Sub chkVisible_LostFocus()
lEnfocaChk = False
End Sub

Private Sub cmdAsigCta_Click()
Dim sOpeCod As String
If tvOpe.Nodes.Count <= 0 Then
   MsgBox "Operación no definida...!", vbInformation, "¡Aviso!"
   Exit Sub
End If
glAceptar = False
sOpeCod = tvOpe.SelectedItem.Tag
frmMntOperacionCta.Inicio sOpeCod, Mid(tvOpe.SelectedItem.Text, 10, 60)
If glAceptar Then
   MuestraDatos sOpeCod
End If
End Sub

Private Sub cmdAsigObj_Click()
Dim sOpeCod As String
If tvOpe.Nodes.Count <= 0 Then
   Exit Sub
End If
sOpeCod = tvOpe.SelectedItem.Tag
frmAsignaObj.Inicio sOpeCod, Mid(tvOpe.SelectedItem.Text, 10, 60), "2"
MuestraObjetos sOpeCod
End Sub

Private Sub CmdBuscar_Click()
Dim clsBuscar As New ClassDescObjeto
Dim K As Long
Set rs = clsOpe.CargaOpeTpo("")
clsBuscar.BuscarDato rs, 0, "Operaciones"
Set clsBuscar = Nothing
If Not rs.EOF Then
   For K = 1 To tvOpe.Nodes.Count
      If tvOpe.Nodes(K).Tag = rs!cOpeCod Then
         tvOpe.Nodes(K).Selected = True
         tvOpe_NodeClick tvOpe.Nodes(K)
         Exit For
      End If
   Next
End If
tvOpe.SetFocus
End Sub

Private Sub cmdEliminar_Click()
On Error GoTo Salida
gsOpeCod = tvOpe.SelectedItem.Tag

If Not tvOpe.SelectedItem.Child Is Nothing Then
    MsgBox "Operación tiene elementos en Niveles inferiores...", vbInformation, "¡Aviso!"
    Exit Sub
End If

If Not rc.EOF Or Not ro.EOF Or Not rd.EOF Then
   If MsgBox(" Operación tiene Cuentas, Objetos o Documentos Asignados. ¿ Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2, "¡Confirmación!") = vbNo Then
      Exit Sub
   End If
Else
   If MsgBox("¿ Esta seguro de eliminar la operación ?", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
      Exit Sub
   End If
End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaMantClasifOperacion
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, 3, Me.Caption & " : Se Elimino la Operacion |CodOpe : " & gsOpeCod & " |Descripcion : " & frmMntOperacionDato.pOpeDesc
            Set objPista = Nothing
            '*******
clsOpe.EliminaOpeTpo gsOpeCod
tvOpe.Nodes.Remove tvOpe.SelectedItem.Index
tvOpe_NodeClick tvOpe.SelectedItem
tvOpe.SetFocus
Exit Sub
Salida:
   MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub cmdImprimir_Click()
Dim sTexto As String
MousePointer = 11
Set oImp = New NContImprimir
sTexto = oImp.ImprimeOperaciones()
MousePointer = 0
Set oImp = Nothing
EnviaPrevio sTexto, "Lista de Operaciones", gnLinPage, False
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaMantClasifOperacion
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", Me.Caption & " : Se Imprimieron las Operaciones "
            Set objPista = Nothing
            '*******
End Sub

Private Sub cmdModificar_Click()
On Error GoTo ErrMod
If tvOpe.Nodes.Count = 0 Then
   Exit Sub
End If
gsOpeCod = tvOpe.SelectedItem.Tag
sOpeDesc = tvOpe.SelectedItem.Text
sOpeDesc = Mid(sOpeDesc, InStr(sOpeDesc, "-") + 2, Len(sOpeDesc))
glAceptar = False
frmMntOperacionDato.Inicio gsOpeCod, sOpeDesc, chkVisible.value, txtOpeNivel, txtOpeGruCod, False
If glAceptar Then
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   clsOpe.ActualizaOpeTpo gsOpeCod, Trim(Replace(frmMntOperacionDato.pOpeDesc, "'", "''")), frmMntOperacionDato.pVisible, frmMntOperacionDato.pOpeNiv, frmMntOperacionDato.pOpeTpo, gsMovNro
   tvOpe.Nodes.Remove tvOpe.SelectedItem.Index
   ActualizaArbolNuevaOperacion gsOpeCod, frmMntOperacionDato.pOpeDesc, frmMntOperacionDato.pOpeNiv
   tvOpe_NodeClick tvOpe.SelectedItem
End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaMantClasifOperacion
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, 2, Me.Caption & " : Se Modifico la Operacion |CodOpe : " & gsOpeCod & " |Descripcion : " & frmMntOperacionDato.pOpeDesc
            Set objPista = Nothing
            '*******
tvOpe.SetFocus
Exit Sub
ErrMod:
   MsgBox TextErr(Err.Description), vbInformation, "!AViso¡"
End Sub

Private Sub cmdNuevo_Click()
Dim nIndOpe As Long
   On Error GoTo NuevoErr
glAceptar = False
frmMntOperacionDato.Inicio "", "", "1", 2, "", True
If glAceptar Then
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   clsOpe.InsertaOpeTpo gsOpeCod, frmMntOperacionDato.pOpeDesc, frmMntOperacionDato.pVisible, frmMntOperacionDato.pOpeNiv, frmMntOperacionDato.pOpeTpo, gsMovNro
   ActualizaArbolNuevaOperacion gsOpeCod, frmMntOperacionDato.pOpeDesc, frmMntOperacionDato.pOpeNiv
End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaMantClasifOperacion
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, 1, Me.Caption & " : Se Agrego la Operacion |CodOpe : " & gsOpeCod & " |Descripcion : " & frmMntOperacionDato.pOpeDesc
            Set objPista = Nothing
            '*******
tvOpe.SetFocus
Exit Sub
NuevoErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub ActualizaArbolNuevaOperacion(psOpeCod As String, psOpeDes As String, pnOpeNiv As Integer)
Dim nItem     As Integer
Dim nPos      As Integer
Dim sOpeKey   As String
Dim sOpePadre As String
Dim nodOpe    As Node
Dim nNivel    As Integer
Dim lnRelacion As Integer

nPos = 0
nNivel = 1
nItem = 1
sOpePadre = ""
SiguienteNiv:
Do While True
    If Not tvOpe.Nodes(nItem).Next Is Nothing Then
        If psOpeCod < Mid(tvOpe.Nodes(nItem).Next, 1, 6) Then
            nPos = nItem
            If nNivel = pnOpeNiv Then
                sOpePadre = tvOpe.Nodes(nItem).Key
                lnRelacion = tvwNext
                Exit Do
            Else
                nNivel = nNivel + 1
                If Not tvOpe.Nodes(nItem).Child Is Nothing Then
                    nItem = tvOpe.Nodes(nItem).Child.Index
                    GoTo SiguienteNiv
                Else
                    'Es el primer Nodo
                    sOpePadre = tvOpe.Nodes(nItem).Key
                    lnRelacion = tvwChild
                    Exit Do
                End If
            End If
        Else
            nItem = tvOpe.Nodes(nItem).Next.Index
        End If
    Else
        If nNivel = pnOpeNiv Then 'Es el ultimo hermano
            sOpePadre = tvOpe.Nodes(nItem).Parent.Key
            lnRelacion = tvwChild
        Else
            sOpePadre = tvOpe.Nodes(nItem).Key
            lnRelacion = tvwChild
        End If
        Exit Do
    End If
Loop

psOpeDes = psOpeCod & " - " & psOpeDes
Select Case pnOpeNiv
    Case "1"
        sOpeKey = "_1" & psOpeCod
        Set nodOpe = tvOpe.Nodes.Add(sOpePadre, lnRelacion, sOpeKey, psOpeDes, "Padre")
        nodOpe.Tag = psOpeCod
    Case "2"
        sOpeKey = "_2" & psOpeCod
        Set nodOpe = tvOpe.Nodes.Add(sOpePadre, lnRelacion, sOpeKey, psOpeDes, "Hijo")
        nodOpe.Tag = psOpeCod
    Case "3"
        sOpeKey = "_3" & psOpeCod
        Set nodOpe = tvOpe.Nodes.Add(sOpePadre, lnRelacion, sOpeKey, psOpeDes, "Hijito")
        nodOpe.Tag = psOpeCod
    Case "4"
        sOpeKey = "_4" & psOpeCod
        Set nodOpe = tvOpe.Nodes.Add(sOpePadre, lnRelacion, sOpeKey, psOpeDes, "Bebe")
        nodOpe.Tag = psOpeCod
End Select
tvOpe.Refresh
tvOpe_NodeClick nodOpe
tvOpe.Nodes(nodOpe.Index).Selected = True
End Sub

Private Sub cmdQuitBien_Click()
gsOpeCod = tvOpe.SelectedItem.Tag
If rb.EOF Or rb.BOF Then
   MsgBox "No existen cuentas para Quitar...", vbCritical, "Error"
   Exit Sub
End If
If MsgBox("¿ Esta seguro que desea quitar la Cuenta Contable de la operación ?", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
   clsOpe.EliminaCtaBS gsOpeCod, rb!cCtaContCod, rb!cBSCod
   rb.Delete adAffectCurrent
End If
tvOpe.SetFocus
End Sub

Private Sub cmdQuitCta_Click()
gsOpeCod = tvOpe.SelectedItem.Tag
If rc.EOF Or rc.BOF Then
   MsgBox "No existen cuentas para Quitar...", vbCritical, "Error"
   Exit Sub
End If
If MsgBox("¿ Esta seguro que desea quitar la Cuenta Contable de la operación ?", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
   clsOpe.EliminaOpeCta gsOpeCod, rc!cOpeCtaOrden, rc!cCtaContCod, rc!cOpeCtaDH
            'ARLO20170208
            Dim lsCodCta, lsCtaOpe, lsOpe As String
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaMantClasifOperacion
            lsCodCta = rc!cCtaContCod
            lsCtaOpe = rc!cCtaContDesc
            lsOpe = tvOpe.SelectedItem.Text
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, 3, Me.Caption & " : Se Quito la |Cta Cod : " & lsCodCta & " | " & lsCtaOpe & " a la Operacion : " & lsOpe
            Set objPista = Nothing
            '*******
   rc.Delete adAffectCurrent
End If
tvOpe.SetFocus
End Sub

Private Sub cmdQuitObj_Click()
If tvOpe.Nodes.Count <= 0 Then
   Exit Sub
End If
gsOpeCod = tvOpe.SelectedItem.Tag
If ro.EOF Or ro.BOF Then
   MsgBox "No existen Objetos para quitar...", vbCritical, "Error"
   Exit Sub
End If
   If MsgBox("¿ Esta seguro que desea quitar el Objeto la operación ?", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
      clsOpe.EliminaOpeObj gsOpeCod, ro!cOpeObjOrden, ro!cObjetoCod
            'ARLO20170208
            Dim lsCodCta, lsCtaOpe, lsOpe As String
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaMantClasifOperacion
            lsCodCta = ro!cObjetoCod
            lsCtaOpe = ro!cObjetoDesc
            lsOpe = tvOpe.SelectedItem.Text
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, 3, Me.Caption & " : Se Quito el Ojbeto : " & lsCodCta & " | " & lsCtaOpe & " a la Operacion : " & lsOpe
            Set objPista = Nothing
            '*******
      ro.Delete adAffectCurrent
      grdObj.SetFocus
  End If
End Sub

Private Sub cmdAsigDoc_Click()
If tvOpe.Nodes.Count <= 0 Then
   Exit Sub
End If
gsOpeCod = tvOpe.SelectedItem.Tag
sOpeDesc = Mid(tvOpe.SelectedItem.Text, 10, 60)
frmMntOperacionDoc.Inicio gsOpeCod, sOpeDesc
MuestraDocumentos gsOpeCod
End Sub

Private Sub cmdQuitDoc_Click()
Dim nDocTpo As Long

If rd.EOF Or rd.BOF Then
   MsgBox "No existen Documentos para quitar...", vbInformation, "Error"
   Exit Sub
End If

gsOpeCod = rd!cOpeCod
nDocTpo = rd!nDocTpo
If MsgBox("¿ Esta seguro que desea quitar el documento de la operación ?", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
   clsOpe.EliminaOpeDoc gsOpeCod, nDocTpo
End If
MuestraDocumentos gsOpeCod
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rs
RSClose rc
RSClose ro
RSClose rd
Set clsOpe = Nothing
frmMdiMain.Enabled = True
End Sub

Private Sub grdCta_GotFocus()
grdCta.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdCta_HeadClick(ByVal ColIndex As Integer)
If Not rc Is Nothing Then
   If Not rc.EOF Then
      rc.Sort = grdCta.Columns(ColIndex).DataField
   End If
End If
End Sub

Private Sub grdCta_LostFocus()
grdCta.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub DesactivaControl()
grdCta.Enabled = False
grdObj.Enabled = False
grdDoc.Enabled = False
cmdAsigCta.Enabled = False
cmdAsigObj.Enabled = False
cmdAsigDoc.Enabled = False
cmdAsigBien.Enabled = False
cmdQuitCta.Enabled = False
cmdQuitObj.Enabled = False
cmdQuitDoc.Enabled = False
cmdQuitBien.Enabled = False
End Sub



Private Sub pageOpe_Click(PreviousTab As Integer)
Dim N As Integer
If tvOpe.Nodes.Count > 0 Then
   gsOpeCod = tvOpe.SelectedItem.Tag
   N = pageOpe.Tab
   Select Case N
       Case 0
            MuestraCuentas (gsOpeCod)
            DesactivaControl
            grdCta.Enabled = True
            cmdAsigCta.Enabled = True
            cmdQuitCta.Enabled = True
       Case 1
            MuestraObjetos (gsOpeCod)
            DesactivaControl
            grdObj.Enabled = True
            cmdAsigObj.Enabled = True
            cmdQuitObj.Enabled = True
       Case 2
            MuestraDocumentos (gsOpeCod)
            DesactivaControl
            grdDoc.Enabled = True
            cmdAsigDoc.Enabled = True
            cmdQuitDoc.Enabled = True
       Case 3
            MuestraOperacionCtaBS (gsOpeCod)
            DesactivaControl
            dgCtaObj.Enabled = True
            cmdAsigBien.Enabled = True
            cmdQuitBien.Enabled = True

   End Select
End If
End Sub
Private Sub Form_Load()
frmMdiMain.Enabled = False
Set clsOpe = New DOperacion
CentraForm Me
LlenaArbolOperacion
MuestraCuentas ""
MuestraObjetos ""
MuestraDocumentos ""

pageOpe.Tab = 0
If lConsulta Then
   cmdNuevo.Visible = False
   cmdModificar.Visible = False
   cmdEliminar.Visible = False
   cmdAsigCta.Visible = False
   cmdAsigDoc.Visible = False
   cmdAsigObj.Visible = False
   cmdQuitCta.Visible = False
   cmdQuitDoc.Visible = False
   cmdQuitObj.Visible = False
   cmdImprimir.Top = cmdNuevo.Top
End If
End Sub
Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub MuestraCuentas(vOpeCod As String)
Set rc = clsOpe.CargaOpeCta(vOpeCod, , , adLockOptimistic)
Set grdCta.DataSource = rc
End Sub

Private Sub MuestraObjetos(vOpeCod As String)
Set ro = clsOpe.CargaOpeObj(vOpeCod, adLockOptimistic)
Set grdObj.DataSource = ro
End Sub

Private Sub MuestraDocumentos(vOpeCod As String)
Set rd = clsOpe.CargaOpeDoc(vOpeCod, , adLockOptimistic)
Set grdDoc.DataSource = rd
End Sub

Sub LlenaArbolOperacion(Optional psOpeCod As String = "")
Dim clsGen As DGeneral
Dim sOperacion As String, sOpeCod As String
Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String
Dim nodOpe As Node
Dim nIndOpe As Long
nIndOpe = 1
Set rs = clsOpe.CargaOpeTpo("", , adLockOptimistic)
tvOpe.Nodes.Clear
Do While Not rs.EOF
    sOpeCod = rs("cOpeCod")
    sOperacion = sOpeCod & " - " & rs("cOpeDesc")
    Select Case rs("nOpeNiv")
        Case "1"
            sOpePadre = "_1" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
            nodOpe.Tag = sOpeCod
        Case "2"
            sOpeHijo = "_2" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
            nodOpe.Tag = sOpeCod
        Case "3"
            sOpeHijito = "_3" & sOpeCod
            Set nodOpe = tvOpe.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
            nodOpe.Tag = sOpeCod
        Case "4"
            Set nodOpe = tvOpe.Nodes.Add(sOpeHijito, tvwChild, "_4" & sOpeCod, sOperacion, "Bebe")
            nodOpe.Tag = sOpeCod
    End Select
    If sOpeCod = psOpeCod And psOpeCod <> "" Then
       nIndOpe = nodOpe.Index
    End If
    rs.MoveNext
Loop
RSClose rs
DoEvents
If tvOpe.Nodes.Count > 0 Then
   tvOpe_NodeClick tvOpe.Nodes(nIndOpe)
   tvOpe.Nodes(nIndOpe).Selected = True
End If
End Sub

Private Sub MuestraDatos(psOpeCod As String)
   Select Case pageOpe.Tab
       Case 0
            MuestraCuentas psOpeCod
       Case 1
            MuestraObjetos psOpeCod
       Case 2
            MuestraDocumentos psOpeCod
       Case 3
            MuestraOperacionCtaBS psOpeCod
   End Select
End Sub

Private Sub MuestraOperacionCtaBS(vOpeCod As String)
Dim oLogFun As New DLogGeneral
Set rb = oLogFun.CargaOperacionCtaBS(vOpeCod)
Set dgCtaObj.DataSource = rb
End Sub


Private Sub tvOpe_NodeClick(ByVal Node As MSComctlLib.Node)
Set rs = clsOpe.CargaOpeTpo(Node.Tag)
If Not rs.EOF Then
   txtOpeGruCod = rs!cOpeGruCod
   txtOpeNivel = rs!nOpeNiv
   If rs!cOpeVisible = "SI" Then
      chkVisible.value = vbChecked
   Else
      chkVisible.value = vbUnchecked
   End If
   Set rs = clsOpe.CargaOpeGru(txtOpeGruCod)
   If Not rs.EOF Then
      txtOpeGruDes = rs!cOpeGruDesc
   End If
End If
RSClose rs
If Left(Node.Tag, 1) = "5" Then
    pageOpe.TabVisible(3) = True
Else
    pageOpe.TabVisible(3) = False
End If
MuestraDatos Node.Tag
End Sub

Private Sub oImp_BarraClose()
oBarra.CloseForm Me
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

