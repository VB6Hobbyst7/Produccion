VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogOCIngBien 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6300
   ClientLeft      =   975
   ClientTop       =   2190
   ClientWidth     =   10425
   Icon            =   "frmLogOCIngBien.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox txtSTotal 
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
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   8940
      TabIndex        =   60
      Top             =   5430
      Width           =   1185
   End
   Begin VB.PictureBox picCuadroSi 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      Picture         =   "frmLogOCIngBien.frx":08CA
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   59
      Top             =   5790
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picCuadroNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3540
      Picture         =   "frmLogOCIngBien.frx":0C0C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   58
      Top             =   5790
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdExaminar 
      Caption         =   "..."
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3030
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3150
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txtCant 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   295
      Left            =   8640
      TabIndex        =   21
      Top             =   2835
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox txtObj 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   295
      Left            =   1725
      TabIndex        =   19
      Top             =   3150
      Visible         =   0   'False
      Width           =   1510
   End
   Begin VB.TextBox txtTotal 
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
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   8940
      TabIndex        =   49
      Top             =   5753
      Width           =   1185
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   4560
      TabIndex        =   27
      Top             =   5910
      Width           =   1230
   End
   Begin VB.PictureBox PicOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10260
      Picture         =   "frmLogOCIngBien.frx":0F4E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   41
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame FrameTipCambio 
      Caption         =   "Tipo de Cambio"
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
      Height          =   645
      Left            =   120
      TabIndex        =   38
      Top             =   5640
      Visible         =   0   'False
      Width           =   3345
      Begin VB.TextBox txtTipCambio 
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   480
         TabIndex        =   25
         Top             =   240
         Width           =   960
      End
      Begin VB.TextBox txtTipCompra 
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   2220
         TabIndex        =   26
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Fijo"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   270
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "Compra"
         Height          =   255
         Left            =   1590
         TabIndex        =   39
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.Frame frameMovDesc 
      Caption         =   "Observaciones"
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
      Height          =   885
      Left            =   120
      TabIndex        =   36
      Top             =   1530
      Width           =   5655
      Begin VB.TextBox txtMovDesc 
         Appearance      =   0  'Flat
         Height          =   585
         Left            =   150
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   210
         Width           =   5355
      End
   End
   Begin VB.Frame frameDestino 
      Caption         =   "Proveedor "
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
      TabIndex        =   29
      Top             =   750
      Width           =   5655
      Begin VB.CommandButton cmdExaCab 
         Caption         =   "..."
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1290
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   300
         Width           =   270
      End
      Begin VB.TextBox txtExaCab 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         TabIndex        =   34
         Top             =   270
         Width           =   285
      End
      Begin VB.TextBox txtProvRuc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   150
         MaxLength       =   20
         TabIndex        =   3
         Tag             =   "txttributario"
         Top             =   270
         Width           =   1185
      End
      Begin VB.TextBox txtProvNom 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         TabIndex        =   5
         Tag             =   "txtnombre"
         Top             =   270
         Width           =   3885
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Movimiento"
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
      Height          =   690
      Left            =   120
      TabIndex        =   31
      Top             =   30
      Width           =   6480
      Begin VB.TextBox txtOpeCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   750
         TabIndex        =   0
         Top             =   240
         Width           =   840
      End
      Begin VB.TextBox txtMovNro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Top             =   240
         Width           =   2940
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   5220
         TabIndex        =   2
         Top             =   255
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label8 
         Caption         =   "Número"
         Height          =   255
         Left            =   150
         TabIndex        =   33
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   4665
         TabIndex        =   32
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   5850
      TabIndex        =   28
      Top             =   5910
      Width           =   1230
   End
   Begin RichTextLib.RichTextBox rtxtAsiento 
      Height          =   315
      Left            =   10230
      TabIndex        =   30
      Top             =   6270
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmLogOCIngBien.frx":1290
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
   Begin TabDlg.SSTab TabDocRef 
      Height          =   1635
      Left            =   5850
      TabIndex        =   7
      Top             =   810
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   2884
      _Version        =   393216
      Style           =   1
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Orden de Com&pra"
      TabPicture(0)   =   "frmLogOCIngBien.frx":1311
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label14"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label19"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtOCPlazo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtOCFecha"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtOCNro"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtOCEntrega"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Guía de &Remisión"
      TabPicture(1)   =   "frmLogOCIngBien.frx":132D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "&Comprobante     "
      TabPicture(2)   =   "frmLogOCIngBien.frx":1349
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox txtOCEntrega 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2550
         TabIndex        =   11
         Top             =   1095
         Width           =   1755
      End
      Begin VB.ComboBox cboDoc 
         Height          =   315
         Left            =   -74370
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   630
         Width           =   3630
      End
      Begin VB.TextBox txtOCNro 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   8
         Top             =   630
         Width           =   1485
      End
      Begin VB.TextBox txtOCFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3225
         TabIndex        =   9
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txtFacSerie 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74355
         MaxLength       =   3
         TabIndex        =   16
         Top             =   1050
         Width           =   495
      End
      Begin VB.TextBox txtFacNro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73860
         MaxLength       =   8
         TabIndex        =   17
         Top             =   1050
         Width           =   1185
      End
      Begin VB.TextBox txtOCPlazo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   10
         Top             =   1095
         Width           =   915
      End
      Begin VB.TextBox txtGRSerie 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74520
         MaxLength       =   3
         TabIndex        =   12
         Top             =   645
         Width           =   375
      End
      Begin VB.TextBox txtGRNro 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74160
         MaxLength       =   8
         TabIndex        =   13
         Top             =   645
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtFacFecha 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   -71850
         TabIndex        =   18
         Top             =   1050
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtGRFecha 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   -71850
         TabIndex        =   14
         Top             =   645
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label19 
         Caption         =   "Entrega"
         Height          =   240
         Left            =   1920
         TabIndex        =   57
         Top             =   1132
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo"
         Height          =   240
         Left            =   -74820
         TabIndex        =   56
         Top             =   660
         Width           =   360
      End
      Begin VB.Label Label4 
         Caption         =   "Nº"
         Height          =   165
         Left            =   -74790
         TabIndex        =   48
         Top             =   1140
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Emisión"
         Height          =   165
         Left            =   -72510
         TabIndex        =   47
         Top             =   1140
         Width           =   705
      End
      Begin VB.Label Label7 
         Caption         =   "Nº"
         Height          =   165
         Left            =   210
         TabIndex        =   46
         Top             =   705
         Width           =   315
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Emisión"
         Height          =   165
         Left            =   2160
         TabIndex        =   45
         Top             =   705
         Width           =   1035
      End
      Begin VB.Label Label14 
         Caption         =   "Plazo"
         Height          =   240
         Left            =   210
         TabIndex        =   44
         Top             =   1132
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Emisión"
         Height          =   165
         Left            =   -72960
         TabIndex        =   43
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label11 
         Caption         =   "Nº"
         Height          =   165
         Left            =   -74790
         TabIndex        =   42
         Top             =   720
         Width           =   315
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000C&
         Height          =   1095
         Left            =   120
         Top             =   450
         Width           =   4245
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H8000000E&
         Height          =   1095
         Left            =   135
         Top             =   465
         Width           =   4245
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000C&
         Height          =   1095
         Left            =   -74880
         Top             =   450
         Width           =   4245
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H8000000E&
         Height          =   1095
         Left            =   -74865
         Top             =   465
         Width           =   4245
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H8000000C&
         Height          =   1095
         Left            =   -74880
         Top             =   450
         Width           =   4245
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H8000000E&
         Height          =   1095
         Left            =   -74865
         Top             =   465
         Width           =   4245
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgDetalle 
      Height          =   2025
      Left            =   120
      TabIndex        =   22
      Top             =   2520
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   3572
      _Version        =   393216
      Rows            =   21
      Cols            =   13
      ForeColorSel    =   -2147483643
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483637
      AllowBigSelection=   0   'False
      Enabled         =   0   'False
      TextStyleFixed  =   3
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   13
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgImp 
      Height          =   915
      Left            =   7320
      TabIndex        =   24
      Top             =   4530
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   1614
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483626
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      GridLinesFixed  =   1
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgObj 
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   4770
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1508
      _Version        =   393216
      Cols            =   5
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483638
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.TextBox txtProvCod 
      Height          =   315
      Left            =   270
      MaxLength       =   20
      TabIndex        =   37
      Tag             =   "txtcodigo"
      Top             =   1020
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sub-Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7860
      TabIndex        =   61
      Top             =   5490
      Width           =   855
   End
   Begin VB.Label Label18 
      Caption         =   "Detalle de Servicio"
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
      Height          =   225
      Left            =   120
      TabIndex        =   55
      Top             =   4590
      Width           =   1785
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impuestos"
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
      Left            =   6360
      TabIndex        =   53
      Top             =   4680
      Width           =   915
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "y/o"
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
      Height          =   225
      Left            =   6360
      TabIndex        =   52
      Top             =   4920
      Width           =   915
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Retenc."
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
      Height          =   165
      Left            =   6390
      TabIndex        =   51
      Top             =   5130
      Width           =   885
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7950
      TabIndex        =   50
      Top             =   5790
      Width           =   615
   End
   Begin VB.Label lblDoc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  NOTA DE INGRESO          Nº 00000000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   675
      Left            =   6675
      TabIndex        =   35
      Top             =   60
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Shape ShapeTotal 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   315
      Left            =   7590
      Top             =   5730
      Width           =   2535
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   6300
      TabIndex        =   54
      Top             =   4530
      Width           =   1005
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   315
      Left            =   7590
      Top             =   5430
      Width           =   2535
   End
   Begin VB.Menu mnuLog 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAtender 
         Caption         =   "&Atender"
      End
      Begin VB.Menu mnuNoAtender 
         Caption         =   "&No Atender"
      End
   End
   Begin VB.Menu mnuObj 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAgregar 
         Caption         =   "&Agregar"
      End
      Begin VB.Menu mnuEliminar 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu mnuGravado 
         Caption         =   "&Gravado"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmLogOCIngBien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String, sSqlObj As String
Dim lTransActiva As Boolean      ' Controla si la transaccion esta activa o no
Dim rs As New ADODB.Recordset    'Rs temporal para lectura de datos
Dim lMN As Boolean, sMoney As String  'Identifican el Simbolo del Tipo de Moneda
Dim nUltimoTxt As String 'Ultimo TextBox seleccionado
Dim lSalir As Boolean
Dim lLlenaObj As Boolean, OK As Boolean, lbBienes As Boolean
Dim sObjCod As String, sObjDesc As String, sObjUnid As String
Dim sCtaCod As String, sCtaDesc As String
Dim sProvCod As String
Dim nTasaIGV As Currency, nVariaIGV As Currency
Dim aCtaCambio(1, 2) As String
Dim lNewProv  As Boolean
Dim sDocDesc  As String
Dim sCtaProvis As String

Dim lnColorBien As Double
Dim lnColorServ As Double
Dim lsOpeCod As String

Public Sub Inicio(LlenaObj As Boolean, OrdenCompra As Boolean, psOpeCod As String, Optional ProvCod As String = "")
    lsOpeCod = psOpeCod
    gcOpeCod = lsOpeCod
    lLlenaObj = LlenaObj
    lbBienes = OrdenCompra
    sProvCod = ProvCod
    Me.Show 1
End Sub

Private Sub ControlesTabDocRef(lGR As Boolean, lFac As Boolean)
txtGRSerie.Enabled = lGR
txtGRNro.Enabled = lGR
txtGRFecha.Enabled = lGR
txtFacSerie.Enabled = lFac
txtFacNro.Enabled = lFac
txtFacFecha.Enabled = lFac
cboDoc.Enabled = lFac
End Sub

Private Sub FormatoImpuesto()
fgImp.ColWidth(0) = 250
fgImp.ColWidth(1) = 750
fgImp.ColWidth(2) = 550
fgImp.ColWidth(3) = 0    'CtaContCod
fgImp.ColWidth(4) = 0    'CtaContDes
fgImp.ColWidth(5) = 0    'D/H
fgImp.ColWidth(6) = 1200
fgImp.ColWidth(7) = 0 'Destino 0/1
fgImp.ColWidth(8) = 0 'Obligatorio, Opcional 1/2
fgImp.ColWidth(9) = 0 'Total Impuesto no Gravado
fgImp.TextMatrix(0, 1) = "Impuesto"
fgImp.TextMatrix(0, 2) = "Tasa"
fgImp.TextMatrix(0, 6) = "Monto"
End Sub

Private Sub FormatoObjeto()
fgObj.TextMatrix(0, 0) = " #"
fgObj.TextMatrix(0, 1) = "Ord"
fgObj.TextMatrix(0, 2) = "Código"
fgObj.TextMatrix(0, 3) = "Descripción"
fgObj.TextMatrix(0, 4) = "SubCta"
fgObj.ColWidth(0) = 350
fgObj.ColWidth(1) = 400
fgObj.ColWidth(2) = 1200
fgObj.ColWidth(3) = 3000
fgObj.ColWidth(4) = 780
fgObj.ColAlignment(1) = 7
fgObj.ColAlignment(2) = 1
End Sub

Private Sub FormatoOrden()
fgDetalle.TextMatrix(0, 0) = "#"
fgDetalle.TextMatrix(0, 1) = "Código"
fgDetalle.TextMatrix(0, 2) = "Descripción"
fgDetalle.TextMatrix(0, 3) = "Unidad"
fgDetalle.TextMatrix(0, 4) = "Cantidad"
fgDetalle.TextMatrix(0, 5) = "P.Unitario"
fgDetalle.TextMatrix(0, 6) = "Saldo"
fgDetalle.TextMatrix(0, 7) = "Sub Total"
fgDetalle.TextMatrix(0, 11) = "Cant.Orden"
fgDetalle.TextMatrix(0, 12) = "Monto Orden"
fgDetalle.ColWidth(0) = 350
fgDetalle.ColWidth(1) = 1500
fgDetalle.ColWidth(2) = 3615
fgDetalle.ColWidth(3) = 900
fgDetalle.ColWidth(4) = 1200
fgDetalle.ColWidth(5) = 1200
fgDetalle.ColWidth(6) = 0
fgDetalle.ColWidth(7) = 1200
fgDetalle.ColWidth(8) = 0     'Cuenta Debe
fgDetalle.ColWidth(9) = 0     'Cuenta Haber - No se Usa
fgDetalle.ColWidth(10) = 0    'B o S
fgDetalle.ColWidth(11) = 1200
fgDetalle.ColWidth(12) = 1200

fgDetalle.ColAlignment(1) = 1
fgDetalle.ColAlignmentFixed(0) = 4
fgDetalle.ColAlignmentFixed(4) = 7
fgDetalle.ColAlignmentFixed(5) = 7
fgDetalle.ColAlignmentFixed(6) = 7
fgDetalle.ColAlignmentFixed(7) = 7

fgDetalle.RowHeight(-1) = 285
End Sub

Private Sub Control_Check()
'If fgImp.TextMatrix(fgImp.Row, 8) = "2" Then
   If fgImp.TextMatrix(fgImp.Row, 0) = "." Then
      Set fgImp.CellPicture = picCuadroNo.Picture
      fgImp.TextMatrix(fgImp.Row, 0) = ""
      fgImp.TextMatrix(fgImp.Row, 6) = ""
   Else
      Set fgImp.CellPicture = picCuadroSi.Picture
      fgImp.TextMatrix(fgImp.Row, 0) = "."
   End If
   CalculaTotal
'End If
End Sub

Private Sub ImprimeNotaIngreso()
Dim n As Integer
Dim sTexto As String
Dim sDesc  As String
  sTexto = "N O T A   D E   I N G R E S O    Nro. " & gcDocNro
  'rtxtAsiento.Text = ImpreCabAsiento(sTexto, False)
  rtxtAsiento.Text = rtxtAsiento.Text & "Proveedor     : " & BON & txtProvNom & BOFF & "    RUC : " & txtProvRuc & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
  sTexto = ""
  If txtOCNro <> "" Then
     sTexto = "Orden de Compra No " & txtOCNro & oImpresora.gPrnSaltoLinea
  End If
  If txtGRNro <> "" Then
     sTexto = sTexto & IIf(sTexto = "", "", Space(16)) & "Guía de Remisión No " & txtGRSerie & "-" & txtGRNro & oImpresora.gPrnSaltoLinea
  End If
  If txtFacNro <> "" Then
     sTexto = sTexto & IIf(sTexto = "", "", Space(16)) & Trim(Left(cboDoc, Len(cboDoc) - 2)) & " No " & txtFacSerie & "-" & txtFacNro & "  " & txtFacFecha & oImpresora.gPrnSaltoLinea
  End If
  If Trim(sTexto) <> "" Then
     rtxtAsiento.Text = rtxtAsiento.Text & "Referencias   : " & BON & sTexto & oImpresora.gPrnSaltoLinea & BOFF
  End If
  rtxtAsiento.Text = rtxtAsiento.Text & ImpreGlosa("Observaciones : ", 80)

sTexto = CON
For n = 1 To fgDetalle.Rows - 1
    sTexto = sTexto & Format(fgDetalle.TextMatrix(n, 0), "00")
    sTexto = sTexto & " " & Left(fgDetalle.TextMatrix(n, 1) + Space(18), 18)
    sTexto = sTexto & " " & Left(fgDetalle.TextMatrix(n, 2) + Space(63), 63)
    sTexto = sTexto & " " & Left(fgDetalle.TextMatrix(n, 3) + Space(10), 6)
    sTexto = sTexto & " " & Right(Space(10) & fgDetalle.TextMatrix(n, 4), 10)
    sTexto = sTexto & " " & Right(Space(12) & fgDetalle.TextMatrix(n, 5), 12)
    sTexto = sTexto & " " & Right(Space(12) & fgDetalle.TextMatrix(n, 7), 12) & oImpresora.gPrnSaltoLinea
    If Len(fgDetalle.TextMatrix(n, 2)) > 63 Then
       sDesc = Mid(fgDetalle.TextMatrix(n, 2), 64, Len(fgDetalle.TextMatrix(n, 2)))
       Do While sDesc <> ""
          Lin1 sTexto, CON & Space(20) & Left(sDesc + Space(63), 63) & COFF
          sDesc = Mid(sDesc, 64, Len(sDesc))
       Loop
    End If
Next
sTexto = sTexto & COFF
rtxtAsiento.Text = rtxtAsiento.Text & ImpreDetLog(sTexto, "Unidad Cantidad   P.Unitario   SubTotal")
rtxtAsiento.Text = rtxtAsiento.Text & oImpresora.gPrnSaltoLinea

rtxtAsiento.Text = rtxtAsiento.Text & BON
rtxtAsiento.Text = rtxtAsiento.Text & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
rtxtAsiento.Text = rtxtAsiento.Text & "      _____________________                        _____________________" & oImpresora.gPrnSaltoLinea
rtxtAsiento.Text = rtxtAsiento.Text & "         Vo Bo ALMACEN                                Vo Bo LOGISTICA   " & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoPagina

rtxtAsiento.Text = ImpreCarEsp(rtxtAsiento.Text)
'frmPrevio.Previo rtxtAsiento, "Documento: Nota de Ingreso", False, gnLinPage
End Sub


Private Function FormaSelect(Optional sObj As String, Optional nNiv As Integer) As String
Dim sText As String
sText = "SELECT d.cObjetoCod, upper(d.cObjetoDesc) as cObjetoDesc, d.nObjetoNiv, c.nCtaObjNiv, b.cCtaContCod, b.cCtaContDesc, e.cUnidadAbrev " _
      & "FROM  " & gcCentralCom & "OpeCta a,  " & gcCentralCom & "CtaCont b,  " & gcCentralCom & "CtaObj c,  " & gcCentralCom & "Objeto d LEFT JOIN Bienes e ON (e.cObjetoCod = d.cObjetoCod) " _
      & "WHERE b.cCtaContCod = a.cCtaContCod AND c.cCtaContCod = b.cCtaContCod AND " _
      & "      ((d.cObjetoCod Like c.cObjetoCod+'%') "
sText = sText & "AND (a.cOpeCod='" & gcOpeCod & "') AND (a.cOpeCtaDH='D')) "
If nNiv > 0 Then
   sText = sText & "and d.nobjetoniv = " & nNiv & " "
End If
FormaSelect = sText & IIf(sObj <> "", "and d.cObjetoCod = '" & sObj & "' ", sObj) _
            & "order by d.cObjetoCod"
End Function

Private Function ValidarProvee(ProvRuc As String) As Boolean
'Dim sSqlProv As String
'Dim rsProv As New ADODB.Recordset
'ValidarProvee = False
'   If Len(Trim(ProvRuc)) = 0 Then
'      Exit Function
'   End If
'   sSqlProv = gcCentralPers & "spBuscaClienteDoc '" & Trim(ProvRuc) & "'"
'   If rsProv.State = adStateOpen Then rsProv.Close: Set rsProv = Nothing
'   rsProv.Open sSqlProv, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
'   If rsProv.EOF Then
'      MsgBox " Proveedor no Encontrado ...! ", vbCritical, "Aviso"
'   Else
'      txtProvNom = rsProv!cNomPers
'      txtProvCod = rsProv!cCodPers
'      lNewProv = False
'      ValidarProvee = True
'   End If
'rsProv.Close: Set rsProv = Nothing
End Function

Private Sub cboDoc_Click()
Dim lvItem As ListItem
Dim nRow As Integer
Dim oCon As DConecta
Set oCon = New DConecta
Dim oMov As DMov
Set oMov = New DMov
oCon.AbreConexion
   fgDetalle.Cols = 13
   fgImp.Clear
   fgImp.TextMatrix(0, 1) = "Impuesto"
   fgImp.TextMatrix(0, 2) = "Tasa"
   fgImp.TextMatrix(0, 6) = "Monto"
   fgImp.Rows = 2

   sSQL = "SELECT a.cCtaContCod, b.cCtaContDesc, c.cImpAbrev, c.nImpTasa, a.cDocImpDH, a.cDocImpOpc, c.cImpDestino " _
        & "FROM   " & gcCentralCom & "DocImpuesto a,  " & gcCentralCom & "CtaCont b,  " & gcCentralCom & "Impuesto c  " _
        & "WHERE  c.cCtaContCod = a.cCtaContCod and b.cCtaContCod = a.cCtaContCod and " _
        & "       a.nDocTpo = '" & Right(cboDoc.Text, 2) & "'"
   Set rs = oCon.CargaRecordSet(sSQL)
   Do While Not rs.EOF
      'Primero adicionamos Columna de Impuesto
      fgDetalle.Cols = fgDetalle.Cols + 1
      fgDetalle.ColWidth(fgDetalle.Cols - 1) = 1200
      fgDetalle.TextMatrix(0, fgDetalle.Cols - 1) = rs!cimpabrev

      AdicionaRow fgImp
      fgImp.Col = 0
      nRow = fgImp.Row
      fgImp.TextMatrix(nRow, 0) = ""
      If rs!cDocImpOpc = "1" Then
         Set fgImp.CellPicture = picCuadroSi.Picture
         fgImp.Text = "."
      Else
         Set fgImp.CellPicture = picCuadroNo.Picture
         fgImp.Text = ""
         fgImp.TextMatrix(fgImp.Row, 6) = ""
      End If
      fgImp.TextMatrix(nRow, 1) = rs!cimpabrev
      fgImp.TextMatrix(nRow, 2) = Format(rs!nImpTasa, gcFormView)
      fgImp.TextMatrix(nRow, 3) = rs!cCtaContCod
      fgImp.TextMatrix(nRow, 4) = rs!cCtaContDesc
      fgImp.TextMatrix(nRow, 5) = rs!cDocImpDH
      fgImp.TextMatrix(nRow, 7) = rs!cImpDestino
      fgImp.TextMatrix(nRow, 8) = rs!cDocImpOpc
      rs.MoveNext
   Loop
   fgImp.Col = 1
   If Right(cboDoc.Text, 2) = "44" Then
      gcDocNro = oMov.GeneraDocNro("44", gMonedaNacional, gsCodUser)
      txtFacSerie.MaxLength = 4
      txtFacSerie = Mid(gcDocNro, 1, 4)
      txtFacSerie.Enabled = False
      txtFacNro = Mid(gcDocNro, 6, 20)
      txtFacNro.Enabled = False
   Else
      txtFacSerie.MaxLength = 3
      txtFacSerie = ""
      txtFacSerie.Enabled = True
      txtFacNro = ""
      txtFacNro.Enabled = True
   End If
   CalculaTotal
End Sub

Private Sub CalculaTotal(Optional lCalcImpuestos As Boolean = True)
Dim n As Integer, m As Integer
Dim nSTot As Currency
Dim nITot As Currency, nImp As Currency
Dim nTot  As Currency
Dim nTotImp As Currency
Dim nTasaImp As Currency
Dim nVV      As Currency
nSTot = 0: nTot = 0
nTotImp = 0: nTasaImp = 0
If fgImp.TextMatrix(1, 1) = "" Then
   lCalcImpuestos = False
End If
For m = 1 To fgImp.Rows - 1
   If fgImp.TextMatrix(m, 0) = "." Then
      nTasaImp = nTasaImp + nVal(IIf(fgImp.TextMatrix(m, 2) = "", 0, fgImp.TextMatrix(m, 2)))
   End If
Next
For m = 1 To fgDetalle.Rows - 1
   nSTot = nSTot + nVal(IIf(fgDetalle.TextMatrix(m, 7) = "", 0, fgDetalle.TextMatrix(m, 7)))
Next

For m = 1 To fgImp.Rows - 1
   nITot = 0
   For n = 1 To fgDetalle.Rows - 1
      If fgImp.TextMatrix(m, 0) = "." And fgDetalle.TextMatrix(n, 1) <> "" Then
         If lCalcImpuestos Then
            nVV = Round(Val(Format(fgDetalle.TextMatrix(n, 7), gcFormDato)) / (1 + (nTasaImp / 100)), 2)
            nImp = Round(nVV * ((Val(Format(fgImp.TextMatrix(m, 2), gcFormDato)) / 100)), 2)
            fgDetalle.TextMatrix(n, m + 12) = Format(nImp, gcFormView)
         Else
            nImp = nVal(fgDetalle.TextMatrix(n, m + 12))
         End If
         nITot = nITot + nImp
      Else
         If lCalcImpuestos Then fgDetalle.TextMatrix(n, m + 12) = ""
      End If
   Next
   If fgImp.TextMatrix(m, 7) = "0" Then
      nTotImp = nTotImp + nITot * IIf(fgImp.TextMatrix(m, 5) = "D", 1, -1)
   End If
   fgImp.TextMatrix(m, 6) = Format(nITot, gcFormView)
   nTot = nTot + nITot * IIf(fgImp.TextMatrix(m, 5) = "D", 1, -1)
Next
txtSTotal = Format(Abs(nTot), gcFormView)
If nTot < 0 Then
   txtSTotal.ForeColor = vbRed
Else
   txtSTotal.ForeColor = vbBlack
End If
txtTotal = Format(nSTot + nTotImp, gcFormView)
End Sub

Private Sub cboDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtFacSerie.SetFocus
End If
End Sub

Private Sub cmdExaCab_Click()
Dim sSqlProv As String
Dim rsProv As New ADODB.Recordset

'frmBuscaCli.Inicia frmLogIngBien, True
lNewProv = False
If Len(txtProvCod.Text) <> 0 Then
   If txtProvRuc.Text = "" Then
      txtProvRuc = "00000000"
   Else
      sSqlProv = "Select cCodPers from  proveedor where cCodPers = '" & txtProvCod & "'"
      'Set rsProv = CargaRecord(sSqlProv)
      If RSVacio(rsProv) Then
         lNewProv = True
      End If
      If rsProv.State = adStateOpen Then rsProv.Close: Set rsProv = Nothing
   End If
   txtMovDesc.SetFocus
End If
End Sub

Private Sub fgDetalle_GotFocus()
txtObj.Visible = False
cmdExaminar.Visible = False
End Sub

Private Sub fgDetalle_Scroll()
    fgDetalle.SetFocus
End Sub

Private Sub fgImp_Click()
If fgImp.TextMatrix(1, 1) = "" Then Exit Sub
If fgImp.Col = 0 Then
   Control_Check
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not OK Then
   If MsgBox(" ¿ Seguro que desea Salir sin grabar Operación ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
      Cancel = 1
      Exit Sub
   End If
End If
End Sub

Private Sub cmdCerrar_Click()
OK = False
lTransActiva = False
Unload Me
End Sub

Private Sub cmdExaminar_Click()
'Dim sSqlO As String
''frmBuscaBien.Inicio 2, 6, "181"
'AbreConexion
'If frmBuscaBien.lOk Then
'   sSqlO = FormaSelect(gaObj(0, 0, UBound(gaObj, 3)))
'   'Set rs = CargaRecord(sSqlO)
'   If RSVacio(rs) Then
'      MsgBox "Objeto no asignado a Operación", vbCritical, "Error"
'      txtObj.SetFocus
'      Exit Sub
'   End If
'   sObjCod = rs!cObjetoCod
'   sObjDesc = rs!cObjetoDesc
'   sObjUnid = rs!cUnidadAbrev
'   sCtaCod = rs!cCtaContCod
'   sCtaDesc = rs!cCtaContDesc
'
'   'ActualizaFG fgDetalle.Row
'   fgDetalle.Enabled = True
'   txtObj.Visible = False
'   cmdExaminar.Visible = False
'   fgDetalle.SetFocus
'   rs.Close
'Else
'   If txtObj.Enabled Then
'      txtObj.SetFocus
'   End If
'End If
End Sub

Private Sub fgDetalle_DblClick()

If fgDetalle.Col > 12 Then
   If fgImp.TextMatrix(fgDetalle.Col - 12, 0) = "." Then
      If fgDetalle.TextMatrix(fgDetalle.Row, 1) = "" Then Exit Sub
      EnfocaTexto txtCant, 0, fgDetalle
   End If
End If

If (fgDetalle.TextMatrix(fgDetalle.Row, 10) = "B" And fgDetalle.Col = 4) Or fgDetalle.Col = 7 Then
   If fgDetalle.TextMatrix(fgDetalle.Row, 1) = "" Then Exit Sub
   EnfocaTexto txtCant, 0, fgDetalle
End If
End Sub

Private Sub fgDetalle_KeyPress(KeyAscii As Integer)
If fgDetalle.Col > 12 Then
   If fgImp.TextMatrix(fgDetalle.Col - 12, 0) = "." Then
      If fgDetalle.TextMatrix(fgDetalle.Row, 1) = "" Then Exit Sub
      EnfocaTexto txtCant, 0, fgDetalle
   End If
End If
If (fgDetalle.TextMatrix(fgDetalle.Row, 10) = "B" And fgDetalle.Col = 4) Or fgDetalle.Col = 7 Then
   If fgDetalle.TextMatrix(fgDetalle.Row, 1) = "" Then Exit Sub

   If InStr("0123456789.", Chr(KeyAscii)) > 0 Then
      EnfocaTexto txtCant, KeyAscii, fgDetalle
   Else
      If KeyAscii = 13 Then EnfocaTexto txtCant, 0, fgDetalle
   End If
End If
End Sub
Private Sub fgDetalle_KeyUp(KeyCode As Integer, Shift As Integer)
If fgDetalle.Col > 12 Then
   If fgImp.TextMatrix(fgDetalle.Col - 12, 0) = "." Then
      KeyUp_Flex fgDetalle, KeyCode, Shift
      CalculaTotal False
   End If
End If
If fgDetalle.Col = 7 Then
   KeyUp_Flex fgDetalle, KeyCode, Shift
   CalculaTotal False
End If
End Sub

Private Sub fgDetalle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   Select Case fgDetalle.Col
          Case 1:
          Case 4: mnuNoAtender.Enabled = False
                  mnuAtender.Enabled = True
                  PopupMenu mnuLog
          Case 5: mnuAtender.Enabled = False
                  mnuNoAtender.Enabled = True
                  PopupMenu mnuLog
   End Select
End If
End Sub

Private Sub Form_Activate()
If lSalir Then
   Unload Me
End If
If Not lLlenaObj Then
   TabDocRef.TabEnabled(0) = False
End If
If Not lbBienes Then
   TabDocRef.TabVisible(1) = False
End If
End Sub

Private Sub Form_Load()
Dim n As Integer, nSaldo As Currency, nCant As Currency
Dim nItem As Integer
Dim oMov As DMov
Set oMov = New DMov
Dim oCon As DConecta
Set oCon = New DConecta
Dim oDoc As DOperacion
Set oDoc = New DOperacion

oCon.AbreConexion

lSalir = False
AbreConexion
If Mid(gcOpeCod, 3, 1) = "2" Then  'Identificación de Tipo de Moneda
   lMN = False
   gsSimbolo = gcME
   If gnTipCambio = 0 Then
      If Not GetTipCambio(gdFecSis, Not gbBitCentral) Then
         lSalir = True
         Exit Sub
      End If
   End If
   FrameTipCambio.Visible = True
   txtTipCambio = Format(gnTipCambio, "##,###,##0.0000")
   txtTipCompra = Format(gnTipCambioC, "##,###,##0.0000")
Else
   lMN = True
   gsSimbolo = gcMN
End If

Set rs = CargaOpeCta(gcOpeCod, "H")
If rs.EOF Then
   MsgBox "Falta definir Cuenta de Provisión en Operación", vbInformation, "¡Aviso!"
   lSalir = True
   Exit Sub
End If
sCtaProvis = rs!cCtaContCod

lnColorBien = "&H00F0FFFF"
lnColorServ = "&H00FFFFC0"

' Defino el Nro de Movimiento
txtOpeCod = gcOpeCod
txtMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
txtFecha = Format(gdFecSis, "dd/mm/yyyy")
If lbBienes Then
   fgDetalle.BackColor = lnColorBien
   fgDetalle.BackColorBkg = lnColorBien
Else
   fgDetalle.BackColor = lnColorServ
   fgDetalle.BackColorBkg = lnColorServ
End If

If sProvCod <> "" Then
   txtProvCod = sProvCod
   sSQL = "select PE.cPersNombre cNomPers, PID.cPersIDnro as cProvRuc From Persona PE Inner Join PersID PID On PE.cPersCod = PID.cPersCod And cPersIDTpo = '2' WHERE PE.cPersCod = '" & txtProvCod & "'"
   'Set rs = oCon.CargaRecordSet(sSQL)
   If RSVacio(rs) Then
      MsgBox "Proveedor no registrado. Por favor verificar", vbCritical, "Error"
      lSalir = True
      Exit Sub
   End If
   txtProvRuc = Trim(rs!cProvRuc)
   txtProvNom = rs!cNomPers
   txtProvRuc.Enabled = False
   cmdExaCab.Enabled = False
   txtOCNro.Tag = gcDocTpo
   txtOCNro = gcDocNro
   txtOCFecha = gdFecha
   txtMovDesc = gsGlosa

   sSQL = " SELECT  mc.nMovNro, mc.nMovItem, len(mo.cBSCod) nMovObjOrden, mc.cCtaContCod, mo.cConsDescripcion, " _
        & " mo.cBSCod, mo.cBSDescripcion, mcd.cDescrip, moc.nMovCant," & IIf(gsSimbolo = gcMN, "mc.nMovImporte ", "me.nMovMeImporte ") & " nMovImporte," _
        & " ISNULL(sust.nMontoAtendido,0) as nMontoAtendido," _
        & " ISNULL(Sust.nCantAtendido, 0) As nCantAtendido" _
        & " FROM MovCta mc " & IIf(gsSimbolo = gcMN, "", " JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem") _
        & " LEFT JOIN ( SELECT nMovNro,nMovItem, mo.cBSCod, cBSDescripcion, cConsDescripcion" _
        & "         FROM MovBS mo" _
        & "         Inner Join BienesServicios o ON o.cBSCod = mo.cBSCod" _
        & "         Inner Join Constante CO On nConsCod = '1019' And nConsValor = nBSUnidad" _
        & "         WHERE nMovNro = '" & gcMovNro & "') mo on mo.nMovNro = mc.nMovNro and mo.nMovItem = mc.nMovItem" _
        & " LEFT JOIN MovCant moc ON moc.nMovNro = mo.nMovNro and moc.nMovItem = mo.nMovItem" _
        & " LEFT JOIN MovCotizacDet mcd ON mcd.nMovNro = mc.nMovNro and mcd.nMovItem = mc.nMovItem" _
        & " LEFT JOIN (SELECT mr.nMovNroRef, mc.nMovItem, SUM(nMov" & IIf(gsSimbolo = gcME, "ME", "") & "Importe) As nMontoAtendido," _
        & "             SUM(moc.nMovCant) As nCantAtendido" _
        & "             FROM MovRef mr" _
        & "             JOIN MovCta mc ON mc.nMovNro = mr.nMovNro " & IIf(gsSimbolo = gcME, " JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
        & "             JOIN MovBS mo ON mo.nMovNro = mc.nMovNro and mo.nMovItem = mc.nMovItem" _
        & "             LEFT JOIN MovCant moc ON moc.nMovNro = mo.nMovNro and moc.nMovItem = mo.nMovItem" _
        & "             JOIN Mov m ON m.nMovNro = mr.nMovNro" _
        & "             WHERE m.nMovEstado = '0' and m.nMovFlag <> '2'" _
        & "             GROUP BY mr.nMovNroRef, mc.nMovItem) Sust" _
        & "             ON Sust.nMovNroRef = mo.nMovNro and Sust.nMovItem = mo.nMovItem" _
        & " WHERE mc.nMovNro = '" & gcMovNro & "' and mc.nMovImporte > 0"

   'sSql = "SELECT mc.cMovNro, mc.cMovItem, mo.cMovObjOrden, mc.cCtaContCod, mo.cObjetoCod, mo.cObjetoDesc, mcd.cDescrip, moc.nMovCant," & IIf(GSSIMBOLO = gcMN, "mc.nMovImporte ", "me.nMovMeImporte ") & " nMovImporte, " _
        & "       ISNULL(sust.nMontoAtendido,0) as nMontoAtendido, ISNULL(sust.nCantAtendido,0) as nCantAtendido " _
        & "FROM MovCta mc " & IIf(GSSIMBOLO = gcMN, "", " JOIN MovMe me ON me.cMovNro = mc.cMovNro and me.cMovItem = mc.cMovItem") _
        & "          LEFT JOIN (SELECT mo.*, o.cObjetoDesc FROM MovObj mo JOIN " & gcCentralCom & "Objeto o  ON o.cObjetoCod = mo.cObjetoCod WHERE not mo.cObjetoCod LIKE '00%' " _
        & "                    ) mo on mo.cMovNro = mc.cMovNro and mo.cMovItem = mc.cMovItem " _
        & "          LEFT JOIN MovCant moc ON moc.cMovNro = mo.cMovNro and moc.cMovItem = mo.cMovItem " _
        & "          LEFT JOIN MovCotizacDet mcd ON mcd.cMovNro = mc.cMovNro and mcd.cMovItem = mc.cMovItem " _
        & "          LEFT JOIN ( SELECT mr.cMovNroRef, mc.cMovItem, mo.cMovObjOrden, SUM(nMov" & IIf(GSSIMBOLO = gcME, "ME", "") & "Importe) as nMontoAtendido, SUM(moc.nMovCant) as nCantAtendido " _
        & "                      FROM MovRef mr JOIN MovCta mc ON mc.cMovNro = mr.cMovNro " & IIf(GSSIMBOLO = gcME, " JOIN MovMe me ON me.cMovNro = mc.cMovNro and me.cMovItem = mc.cMovItem ", "") _
        & "                                     JOIN MovObj mo ON mo.cMovNro = mc.cMovNro and mo.cMovItem = mc.cMovItem " _
        & "                                LEFT JOIN MovCant moc ON moc.cMovNro = mo.cMovNro and moc.cMovItem = mo.cMovItem " _
        & "                                     JOIN Mov m ON m.cMovNro = mr.cMovNro " _
        & "                      WHERE m.cMovEstado = '0' and m.cMovFlag <> 'X' " _
        & "                      GROUP BY mr.cMovNroRef, mc.cMovItem, mo.cMovObjOrden " _
        & "                    ) Sust ON Sust.cMovNroRef = mo.cMovNro and Sust.cMovItem = mo.cMovItem and Sust.cMovObjOrden = mo.cMovObjOrden " _
        & "WHERE mc.cMovNro = '" & gcMovNro & "' and mc.nMovImporte > 0"

   Set rs = oCon.CargaRecordSet(sSQL)
   n = 0
   Do While Not rs.EOF
      n = n + 1
      fgDetalle.TextMatrix(n, 0) = n
      fgDetalle.TextMatrix(n, 7) = Format(rs!nMovImporte - rs!nMontoAtendido, gcFormView)
      fgDetalle.TextMatrix(n, 8) = rs!cCtaContCod
      If Not IsNull(rs!cDescrip) Then
         fgDetalle.TextMatrix(n, 2) = rs!cDescrip
      Else
         fgDetalle.TextMatrix(n, 2) = rs!cObjetoDesc
      End If
      If Not IsNull(rs!nMovCant) Then
         fgDetalle.TextMatrix(n, 1) = rs!cBSCod
         fgDetalle.TextMatrix(n, 3) = rs!cConsDescripcion

         fgDetalle.TextMatrix(n, 4) = rs!nMovCant - rs!nCantAtendido
         fgDetalle.TextMatrix(n, 11) = rs!nMovCant
         fgDetalle.TextMatrix(n, 12) = Format(rs!nMovImporte, gcFormView)

         If rs!nMovCant <> 0 Then
            fgDetalle.TextMatrix(n, 5) = Format(Round(rs!nMovImporte / rs!nMovCant, 2), gcFormView)
         End If
         fgDetalle.TextMatrix(n, 10) = "B"
         FlexBackColor fgDetalle, n, lnColorBien
         rs.MoveNext
      Else
         nItem = rs!nMovItem
         fgDetalle.TextMatrix(n, 1) = rs!cCtaContCod
         fgDetalle.TextMatrix(n, 10) = "S"
         FlexBackColor fgDetalle, n, lnColorServ
         Do While Val(rs!nMovItem) = nItem
            If Not IsNull(rs!cBSCod) Then
               AdicionaRow fgObj
               fgObj.TextMatrix(fgObj.Row, 0) = n
               fgObj.TextMatrix(fgObj.Row, 1) = rs!nMovItem
               fgObj.TextMatrix(fgObj.Row, 2) = rs!cCtaContCod
               fgObj.TextMatrix(fgObj.Row, 3) = rs!cObjetoDesc
            End If
            rs.MoveNext
            If rs.EOF Then
               Exit Do
            End If
         Loop
      End If
   Loop
   sSQL = "SELECT dMovPlazo, cMovLugarEntrega FROM MovCotizac WHERE NMovNro = '" & gcMovNro & "'"
   Set rs = oCon.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      txtOCPlazo = Format(rs!dMovPlazo, "dd/mm/yyyy")
      txtOCEntrega = rs!cMovLugarEntrega
   End If
   RSClose rs
   fgDetalle.Enabled = True
End If

Me.Caption = gcOpeDesc
FormatoOrden
FormatoObjeto
FormatoImpuesto
If lbBienes Then
   Set rs = oDoc.CargaOpeDoc(gcOpeCod, "2")
   If RSVacio(rs) Then
      MsgBox "No se definió Documento NOTA DE INGRESO en Operación. Por favor Consultar con Sistemas...!", vbInformation, "¡Aviso!"
      lSalir = True
      Exit Sub
   Else
      gcDocTpo = rs!nDocTpo
      gcDocNro = oMov.GeneraDocNro(CInt(gcDocTpo), gMonedaExtranjera, Year(gdFecSis))
      sDocDesc = rs!cDocDesc
   End If

   'GUIA DE REMISION
   Set rs = oDoc.CargaOpeDoc(gcOpeCod, "1")
   If RSVacio(rs) Then
      MsgBox "No se definió Documento GUIA DE REMISION en Operación. Por favor Consultar con Sistemas...!", vbInformation, "¡Aviso!"
      lSalir = True
      Exit Sub
   Else
      txtGRNro.Tag = rs!nDocTpo
   End If
   lblDoc.Caption = "   " & UCase(sDocDesc) & "       Nº " & gcDocNro
Else
   TabDocRef.TabCaption(0) = "Orden de Servicio"
   lblDoc.Visible = False
End If
fgObj.BackColor = lnColorServ
'Tipos de Comprobantes de Pago
Set rs = oDoc.CargaOpeDoc(gcOpeCod, "3")
Do While Not rs.EOF
   cboDoc.AddItem (rs!cDocDesc & Space(100) & rs!nDocTpo)
   rs.MoveNext
Loop
If cboDoc.ListCount = 1 Then
   cboDoc.ListIndex = 0
End If
RSClose rs
End Sub

Private Sub TabDocRef_Click(PreviousTab As Integer)
Select Case TabDocRef.Tab
       Case 0
            If txtOCNro.Enabled Then
               txtOCNro.SetFocus
            End If
            ControlesTabDocRef False, False
       Case 1
            ControlesTabDocRef True, False
            txtGRSerie.SetFocus
       Case 2
            ControlesTabDocRef False, True
            cboDoc.SetFocus
End Select
End Sub

Private Sub txtFacFecha_GotFocus()
      txtFacFecha.SelStart = 0
      txtFacFecha.SelLength = Len(txtFacFecha)
End Sub

Private Sub txtFacFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   'If Not IsDate(txtFacFecha.Text) Then
   If ValidaFecha(txtFacFecha) <> "" Then
      MsgBox " Fecha no Válida... ", vbCritical, "Error"
      txtFacFecha.SelStart = 0
      txtFacFecha.SelLength = Len(txtFacFecha)
   Else
      If gsSimbolo = gcME Then
         GetTipCambio CDate(txtFacFecha), Not gbBitCentral
         txtTipCompra = Format(gnTipCambioC, "###,###,##0.0000")
         gnTipCambio = Val(Format(txtTipCambio, gcFormDato))
      End If
      fgDetalle.SetFocus
   End If
End If
End Sub

Private Sub txtFacFecha_Validate(Cancel As Boolean)
If ValidaFecha(txtFacFecha) <> "" Then
   MsgBox " Fecha no Válida... ", vbCritical, "Error"
   txtFacFecha.SelStart = 0
   txtFacFecha.SelLength = Len(txtFacFecha)
   Cancel = True
Else
   GetTipCambio CDate(txtFacFecha), Not gbBitCentral
   txtTipCompra = Format(gnTipCambioC, "###,###,##0.000")
End If
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   If ValidaFecha(txtFecha.Text) <> "" Then
'      MsgBox " Fecha no Válida... ", vbInformation, "Aviso"
'      txtFecha.SelStart = 0
'      txtFecha.SelLength = Len(txtFecha)
'   Else
'      txtMovNro = GeneraMovNro(, , txtFecha)
'      If txtProvRuc.Enabled Then
'         txtProvRuc.SetFocus
'      Else
'         txtMovDesc.SetFocus
'      End If
'   End If
'End If
End Sub

Private Sub txtFecha_Validate(Cancel As Boolean)
Dim oMov As DMov
Set oMov = New DMov
If ValidaFecha(txtFecha.Text) <> "" Then
   MsgBox " Fecha no Válida ", vbInformation, "Aviso"
   Cancel = True
Else
   txtMovNro = oMov.GeneraMovNro(CDate(txtFecha.Text), gsCodAge, gsCodUser)
End If
End Sub

Private Sub txtGRFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtGRFecha.Text) <> "" Then
      MsgBox " Fecha no Válida... ", vbCritical, "Error"
      txtGRFecha.SelStart = 0
      txtGRFecha.SelLength = Len(txtGRFecha)
   Else
      TabDocRef.Tab = 2
   End If
End If
End Sub

Private Sub txtGRNro_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtGRNro = Right(String(8, "0") & txtGRNro, 8)
   txtGRFecha.SetFocus
End If
End Sub
Private Sub txtFacNro_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtFacNro = Right(String(8, "0") & txtFacNro, 8)
   txtFacFecha.SetFocus
End If
End Sub

Private Sub txtGRSerie_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtGRSerie = Right(String(3, "0") & txtGRSerie, 3)
   txtGRNro.SetFocus
End If
End Sub
Private Sub txtFacSerie_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtFacSerie = Right(String(3, "0") & txtFacSerie, 3)
   txtFacNro.SetFocus
End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   Select Case TabDocRef.Tab
          Case 0: TabDocRef.Tab = 2
          Case 1: txtGRSerie.SetFocus
          Case 2: cboDoc.SetFocus
   End Select
End If
End Sub

Private Sub txtObj_GotFocus()
cmdExaminar.Visible = True
cmdExaminar.Top = txtObj.Top + 15
cmdExaminar.Left = txtObj.Left + txtObj.Width - cmdExaminar.Width
End Sub

Private Sub txtProvRuc_GotFocus()
fEnfoque txtProvRuc
End Sub
Private Sub txtProvRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidarProvee(txtProvRuc) Then
      txtMovDesc.SetFocus
   End If
End If
End Sub

Private Sub txtProvRuc_Validate(Cancel As Boolean)
If Not ValidarProvee(txtProvRuc) Then
   Cancel = True
End If
End Sub

Private Sub txtObj_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
   txtObj_KeyPress 13
   SendKeys IIf(KeyCode = 38, "{Up}", "{Down}"), True
End If
End Sub

Private Sub txtObj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaObj Then
      'ActualizaFG fgDetalle.Row
      fgDetalle.Enabled = True
      txtObj.Visible = False
      cmdExaminar.Visible = False
      fgDetalle.SetFocus
   End If
End If
End Sub
Private Function ValidaObj() As Boolean
ValidaObj = False
If Len(txtObj) = 0 Then
   txtObj.Visible = False
   cmdExaminar.Visible = False
   EliminaRow fgDetalle, fgDetalle.Row
   Exit Function
End If
sSQL = FormaSelect(txtObj, 6)
'Set rs = CargaRecord(sSQL)
If Not RSVacio(rs) Then
   sObjCod = rs!cObjetoCod
   sObjDesc = rs!cObjetoDesc
   sObjUnid = rs!cUnidadAbrev
   sCtaCod = rs!cCtaContCod
   sCtaDesc = rs!cCtaContDesc
Else
   MsgBox "Objeto no encontrado...!", vbCritical, "Error de Búsqueda"
   Exit Function
End If
ValidaObj = True
End Function

Private Sub txtCant_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCant, KeyAscii, 12, 2)
If KeyAscii = 13 Then
   If fgDetalle.Col = 4 Then
      If nVal(txtCant) > nVal(fgDetalle.TextMatrix(txtCant.Tag, 11)) Then
         MsgBox "Cantidad Atendida no puede mayor que Cantidad Solicitada", vbInformation, "¡Aviso!"
         Exit Sub
      End If
      fgDetalle.TextMatrix(txtCant.Tag, 7) = Format(Round(Val(txtCant) * Val(Format(fgDetalle.TextMatrix(fgDetalle.Row, 5), "#0.00")), 2), gcFormView)
   End If
   If fgDetalle.Col = 7 And nVal(IIf(fgDetalle.TextMatrix(txtCant.Tag, 4) = "", "0", fgDetalle.TextMatrix(txtCant.Tag, 4))) <> 0 Then
      If nVal(txtCant) > nVal(fgDetalle.TextMatrix(txtCant.Tag, 12)) Then
         MsgBox "Monto no puede ser mayor a Monto pactado", vbInformation, "¡Aviso!"
         Exit Sub
      End If
      fgDetalle.TextMatrix(txtCant.Tag, 5) = Format(Round(Val(txtCant) / Val(Format(fgDetalle.TextMatrix(fgDetalle.Row, 4), "#0.00")), 2), gcFormView)
   End If
   fgDetalle.Text = Format(txtCant.Text, gcFormView)
   txtCant.Visible = False
   fgDetalle.SetFocus
   CalculaTotal False
End If
End Sub

Private Sub txtCant_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
   txtCant_KeyPress 13
   SendKeys IIf(KeyCode = 38, "{Up}", "{Down}"), True
End If
End Sub

Private Sub txtCant_LostFocus()
txtCant.Text = ""
txtCant.Visible = False
End Sub
Private Sub txtCant_Validate(Cancel As Boolean)
'If Val(txtCant) = 0 Then
'   MsgBox "Importe debe ser mayor que Cero...!", vbCritical, "Aviso"
'   Cancel = True
'End If
End Sub

Private Sub cmdAceptar_Click()
    Dim n As Integer 'Contador
    Dim nItem As Integer, nCol  As Integer
    Dim nObj  As Integer
    Dim sTexto As String, lOk As Boolean
    Dim sMovNro As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oMov As DMov
    Set oMov = New DMov
    Dim lnMovNro As Long

    On Error GoTo ErrAceptar


    'ImprimeNotaIngreso
    If txtFacSerie = "" Or txtFacNro = "" Or Not ValidaFecha(txtFacFecha) = "" Then
       MsgBox "Faltan datos del Comprobante del Proveedor. Por favor verificar...!", vbInformation, "¡Aviso!"
       txtFacSerie.SetFocus
       Exit Sub
    End If

    'Primero Salvamos MovNro para hacer referencia cuando es atención de Orden de Compra
    sMovNro = gcMovNro
    gcMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    gsGlosa = txtMovDesc
    gdFecha = txtFecha

If MsgBox(" ¿ Seguro de grabar Operación ? ", vbOKCancel, "Aviso de Confirmación") = vbCancel Then

    Exit Sub
End If

' Iniciamos transaccion
If lTransActiva Then
   lTransActiva = False
End If
oMov.BeginTrans
lTransActiva = True
'Grabamos en Mov

oMov.InsertaMov gcMovNro, gcOpeCod, Me.txtMovDesc.Text
lnMovNro = oMov.GetnMovNro(gcMovNro)

If lbBienes Then
   'Grabamos Nota de Ingreso
   'gcDocNro = GeneraDocNro(CInt(gcDocTpo), gMonedaNacional, Year(gdFecSis))
   'oMov.InsertaMovDoc lnMovNro, CInt(gcDocTpo), gcDocNro, Format(CDate(Me.txtFecha.Text), gsFormatoFecha)

   ' Actualizamos Nro de Nota de Ingreso
   lblDoc.Caption = UCase(sDocDesc) & " Nº " & gcDocNro

   'Ahora la Guía de Remisión
   If txtGRNro <> "" And ValidaFecha(txtGRFecha) = "" Then
      oMov.InsertaMovDoc lnMovNro, CInt(Me.txtGRNro.Tag), IIf(txtGRSerie <> "", txtGRSerie & "-", "") & txtGRNro, Format(CDate(Me.txtGRFecha.Text), gsFormatoFecha)
      sSQL = "INSERT INTO MovDoc (cMovNro, cDocTpo, cDocNro, dDocFecha) VALUES('" & gcMovNro & "', '" & Me.txtGRNro.Tag & "', '" _
           & IIf(txtGRSerie <> "", txtGRSerie & "-", "") & txtGRNro & "', '" & Format(CDate(txtGRFecha), gcFormatoFecha) & "') "
   End If
End If

'Por ultimo, el Comprobante del Proveedor
If txtFacNro <> "" Then
    oMov.InsertaMovDoc lnMovNro, CInt(Right(cboDoc.Text, 2)), IIf(txtFacSerie <> "", txtFacSerie & "-", "") & txtFacNro, Format(CDate(txtFacFecha), gcFormatoFecha)
   sSQL = " INSERT INTO MovDoc (cMovNro, cDocTpo, cDocNro, dDocFecha) VALUES('" & gcMovNro & "', '" & Right(cboDoc.Text, 2) & "', '" _
        & IIf(txtFacSerie <> "", txtFacSerie & "-", "") & txtFacNro & "', '" & Format(CDate(txtFacFecha), gcFormatoFecha) & "') "
End If

' Grabamos en MovObj y MovCant
nItem = 0
For n = 1 To fgDetalle.Rows - 1
   If Len(fgDetalle.TextMatrix(n, 1)) > 0 Then
      nItem = nItem + 1
      If Val(Format(fgDetalle.TextMatrix(n, 7), gcFormDato)) > 0 Then
         If fgDetalle.TextMatrix(n, 10) = "B" Then
            oMov.InsertaMovBS lnMovNro, nItem, Len(fgDetalle.TextMatrix(n, 1)), fgDetalle.TextMatrix(n, 1)
            'GrabaMovObj Format(nItem, "000"), fgDetalle.TextMatrix(N, 1), "1"
         Else
            nObj = 1
            For nCol = 1 To fgObj.Rows - 1
               If fgObj.TextMatrix(nCol, 0) = fgDetalle.TextMatrix(n, 0) Then
                  oMov.InsertaMovObj lnMovNro, nItem, fgObj.TextMatrix(nCol, 2), fgObj.TextMatrix(nCol, 1)
                  nObj = Val(fgObj.TextMatrix(nCol, 1)) + 1
               End If
            Next
            oMov.InsertaMovGasto lnMovNro, txtProvCod, Format(nObj, "0")
         End If
         oMov.InsertaMovCta lnMovNro, nItem, fgDetalle.TextMatrix(n, 8), nVal(fgDetalle.TextMatrix(n, 7))
      End If
      'Si cantidad > 0 then
      If nVal(fgDetalle.TextMatrix(n, 4)) <> 0 Then
           'Actualizamos los Stocks
         oMov.InsertaMovCant lnMovNro, nItem, nVal(fgDetalle.TextMatrix(n, 4))
      End If
   End If
Next

'Actualizamos los Valores de Impuestos cobrados
gnImporte = 0
For nCol = 1 To fgImp.Rows - 1
   If fgImp.TextMatrix(nCol, 0) = "." And fgImp.TextMatrix(nCol, 7) = "1" Then
      For n = 1 To fgDetalle.Rows - 1
         If fgDetalle.TextMatrix(n, 1) <> "" And nVal(IIf(fgDetalle.TextMatrix(n, 7) = "", "0", fgDetalle.TextMatrix(n, 7))) > 0 Then
            oMov.InsertaMovOtrosItem lnMovNro, n, fgImp.TextMatrix(nCol, 3), fgDetalle.TextMatrix(n, nCol + 12), "003"
         End If
      Next
   End If
   'Generamos el Asiento Complementario de Impuestos o Retenciones
   If fgImp.TextMatrix(nCol, 0) = "." And fgImp.TextMatrix(nCol, 7) = "0" Then
      nItem = nItem + 1
      gnImporte = nVal(fgImp.TextMatrix(nCol, 6)) * IIf(fgImp.TextMatrix(nCol, 5) = "D", 1, -1)
      oMov.InsertaMovCta lnMovNro, nItem, fgImp.TextMatrix(nCol, 3), gnImporte
   End If
Next
'Grabamos la Provisión
If nVal(txtTotal) <> 0 Then
   nItem = nItem + 1
   gnImporte = nVal(txtTotal)
   oMov.InsertaMovCta lnMovNro, nItem, sCtaProvis, gnImporte * -1
   oMov.InsertaMovGasto lnMovNro, txtProvCod, ""
End If
'Grabamos MovRef Referencias a la Orden de Compra ...si existe
'Primero referenciamos a la Orden de Compra
If txtOCNro <> "" Then
   oMov.InsertaMovRef lnMovNro, sMovNro
End If

'Grabamos el Tipo de Cambio de Compra
If Not lMN Then
    oMov.GeneraMovME lnMovNro, gcMovNro
End If
oMov.CommitTrans

gnSaldo = gnSaldo - nVal(txtTotal)

lTransActiva = False
   OK = True
   If lbBienes Then
      'ImprimeNotaIngreso
   Else
      If Right(cboDoc, 2) = "44" Then
         ImprimeReciboEgreso
      End If
   End If
   Unload Me
Exit Sub
ErrAceptar:
    If lTransActiva Then
        oMov.RollbackTrans
    End If
    lTransActiva = False
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub ImprimeReciboEgreso()
    Dim sTexto As String
    Dim nAncho As Integer, n As Integer
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    rtxtAsiento.Text = ""
    nAncho = gnColPage
    gcMovNro = txtMovNro
    gsGlosa = txtMovDesc
    sTexto = ""
    Lin1 sTexto, ImpreCabAsiento(" R E C I B O   D E   E G R E S O  Nro. " & gcDocNro, gdFecSis, gcEmpresaLogo, gcOpeCod, Me.txtMovNro.Text)
    Lin1 sTexto, "  Persona      : " & BON & txtProvCod & "  " & ImpreCarEsp(txtProvNom) & BOFF
    Lin1 sTexto, "  Cargo        : "
    Lin1 sTexto, "  Importe      : " & BON & ConvNumLet(nVal(txtTotal), False) & BOFF
    Lin1 sTexto, ImpreGlosa("  Concepto     : ", 80)
    Lin1 sTexto, "", 6
    Lin1 sTexto, BON & Space(15) & "_________________________                ______________________"
    Lin1 sTexto, Space(15) & CON & Centra(ImpreCarEsp(Mid(txtProvNom, 1, 40)), 40) & COFF & "                    Vo Bo JEFE AREA " & BOFF
    rtxtAsiento.Text = sTexto & String(3, oImpresora.gPrnSaltoLinea)
    If Dir(App.path & "\SPOOLER", vbDirectory) <> "" Then
       rtxtAsiento.SaveFile App.path & "\SPOOLER\ReciboEgreso.txt"
    End If
    oPrevio.Show rtxtAsiento.Text, "Recibo de Egreso", False, gnLinPage / 2
End Sub

Public Property Get lOk() As Boolean
    lOk = OK
End Property

Public Property Let lOk(ByVal vNewValue As Boolean)
    OK = vNewValue
End Property
