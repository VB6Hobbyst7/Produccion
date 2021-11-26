VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPlaOCIngBien 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Presupuestal"
   ClientHeight    =   6300
   ClientLeft      =   450
   ClientTop       =   1740
   ClientWidth     =   10830
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraComentario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Comentario"
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
      Height          =   780
      Left            =   5985
      TabIndex        =   53
      Top             =   4860
      Width           =   4755
      Begin VB.TextBox txtComentario 
         Appearance      =   0  'Flat
         Height          =   450
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Top             =   255
         Width           =   4545
      End
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "&Rechazar"
      Height          =   360
      Left            =   6945
      TabIndex        =   52
      Top             =   5895
      Width           =   1230
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
      ForeColor       =   &H000040C0&
      Height          =   300
      Left            =   6810
      TabIndex        =   51
      Top             =   4530
      Width           =   1185
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   9525
      TabIndex        =   22
      Top             =   5895
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
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
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   45
      TabIndex        =   16
      Top             =   30
      Width           =   7245
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
         TabIndex        =   18
         Top             =   240
         Width           =   2940
      End
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
         TabIndex        =   17
         Top             =   240
         Width           =   840
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   5325
         TabIndex        =   19
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   4770
         TabIndex        =   21
         Top             =   277
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "Número"
         Height          =   255
         Left            =   150
         TabIndex        =   20
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.Frame frameDestino 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      ForeColor       =   &H00800000&
      Height          =   750
      Left            =   45
      TabIndex        =   11
      Top             =   750
      Width           =   5655
      Begin VB.TextBox txtProvNom 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         TabIndex        =   15
         Tag             =   "txtnombre"
         Top             =   270
         Width           =   3885
      End
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
         TabIndex        =   12
         Top             =   300
         Width           =   270
      End
      Begin VB.TextBox txtExaCab 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         TabIndex        =   13
         Top             =   270
         Width           =   285
      End
      Begin VB.TextBox txtProvRuc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   150
         MaxLength       =   20
         TabIndex        =   14
         Tag             =   "txttributario"
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame frameMovDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      ForeColor       =   &H00800000&
      Height          =   885
      Left            =   45
      TabIndex        =   9
      Top             =   1530
      Width           =   5655
      Begin VB.TextBox txtMovDesc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   585
         Left            =   150
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   210
         Width           =   5355
      End
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
      Left            =   45
      TabIndex        =   4
      Top             =   5640
      Visible         =   0   'False
      Width           =   3345
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
         TabIndex        =   6
         Top             =   240
         Width           =   960
      End
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
         TabIndex        =   5
         Top             =   255
         Width           =   960
      End
      Begin VB.Label Label5 
         Caption         =   "Compra"
         Height          =   255
         Left            =   1590
         TabIndex        =   8
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Fijo"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   435
      End
   End
   Begin VB.PictureBox PicOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10185
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   8235
      TabIndex        =   2
      Top             =   5895
      Width           =   1230
   End
   Begin VB.PictureBox picCuadroNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3465
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   5790
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picCuadroSi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3765
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   5790
      Visible         =   0   'False
      Width           =   255
   End
   Begin RichTextLib.RichTextBox rtxtAsiento 
      Height          =   255
      Left            =   4305
      TabIndex        =   23
      Top             =   5910
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   0   'False
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmPlaOCIngBien.frx":0000
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
      Left            =   5775
      TabIndex        =   24
      Top             =   810
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   2884
      _Version        =   393216
      Style           =   1
      Tab             =   2
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
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label19"
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(3)=   "Label14"
      Tab(0).Control(4)=   "Shape1"
      Tab(0).Control(5)=   "Shape4"
      Tab(0).Control(6)=   "txtOCEntrega"
      Tab(0).Control(7)=   "txtOCNro"
      Tab(0).Control(8)=   "txtOCFecha"
      Tab(0).Control(9)=   "txtOCPlazo"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Guía de &Remisión"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(2)=   "Shape2"
      Tab(1).Control(3)=   "Shape3"
      Tab(1).Control(4)=   "txtGRFecha"
      Tab(1).Control(5)=   "txtGRSerie"
      Tab(1).Control(6)=   "txtGRNro"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "&Comprobante     "
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label9"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Shape5"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Shape6"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtFacFecha"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cboDoc"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtFacSerie"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtFacNro"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.TextBox txtGRNro 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74040
         MaxLength       =   8
         TabIndex        =   33
         Top             =   630
         Width           =   1350
      End
      Begin VB.TextBox txtGRSerie 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74475
         MaxLength       =   3
         TabIndex        =   32
         Top             =   630
         Width           =   375
      End
      Begin VB.TextBox txtOCPlazo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74280
         TabIndex        =   31
         Top             =   1103
         Width           =   1350
      End
      Begin VB.TextBox txtFacNro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1170
         MaxLength       =   8
         TabIndex        =   30
         Top             =   1050
         Width           =   1590
      End
      Begin VB.TextBox txtFacSerie 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   630
         MaxLength       =   3
         TabIndex        =   29
         Top             =   1050
         Width           =   495
      End
      Begin VB.TextBox txtOCFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -71505
         TabIndex        =   28
         Top             =   630
         Width           =   1200
      End
      Begin VB.TextBox txtOCNro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74520
         MaxLength       =   13
         TabIndex        =   27
         Top             =   630
         Width           =   1575
      End
      Begin VB.ComboBox cboDoc 
         Height          =   315
         Left            =   630
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   630
         Width           =   4050
      End
      Begin VB.TextBox txtOCEntrega 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72090
         TabIndex        =   25
         Top             =   1103
         Width           =   1785
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
         Left            =   3525
         TabIndex        =   34
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
         Left            =   -71445
         TabIndex        =   35
         Top             =   630
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
      Begin VB.Shape Shape6 
         BorderColor     =   &H8000000E&
         Height          =   1095
         Left            =   135
         Top             =   465
         Width           =   4785
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H8000000C&
         Height          =   1095
         Left            =   120
         Top             =   450
         Width           =   4770
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H8000000E&
         Height          =   1095
         Left            =   -74865
         Top             =   465
         Width           =   4755
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000C&
         Height          =   1095
         Left            =   -74880
         Top             =   450
         Width           =   4770
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H8000000E&
         Height          =   1095
         Left            =   -74850
         Top             =   450
         Width           =   4665
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000C&
         Height          =   1095
         Left            =   -74865
         Top             =   435
         Width           =   4680
      End
      Begin VB.Label Label11 
         Caption         =   "Nº"
         Height          =   165
         Left            =   -74790
         TabIndex        =   44
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Emisión"
         Height          =   165
         Left            =   -72570
         TabIndex        =   43
         Top             =   705
         Width           =   1035
      End
      Begin VB.Label Label14 
         Caption         =   "Plazo"
         Height          =   240
         Left            =   -74790
         TabIndex        =   42
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Emisión"
         Height          =   165
         Left            =   -72750
         TabIndex        =   41
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label7 
         Caption         =   "Nº"
         Height          =   165
         Left            =   -74790
         TabIndex        =   40
         Top             =   690
         Width           =   315
      End
      Begin VB.Label Label6 
         Caption         =   "Emisión"
         Height          =   165
         Left            =   2865
         TabIndex        =   39
         Top             =   1125
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "Nº"
         Height          =   165
         Left            =   210
         TabIndex        =   38
         Top             =   1125
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo"
         Height          =   240
         Left            =   180
         TabIndex        =   37
         Top             =   660
         Width           =   360
      End
      Begin VB.Label Label19 
         Caption         =   "Entrega"
         Height          =   240
         Left            =   -72750
         TabIndex        =   36
         Top             =   1170
         Width           =   1245
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgDetalle 
      Height          =   2025
      Left            =   45
      TabIndex        =   45
      Top             =   2535
      Width           =   10695
      _ExtentX        =   18865
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
      AllowUserResizing=   1
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   13
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgObj 
      Height          =   885
      Left            =   60
      TabIndex        =   47
      Top             =   4770
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   1561
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgImp 
      Height          =   60
      Left            =   7035
      TabIndex        =   46
      Top             =   4530
      Visible         =   0   'False
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   106
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483626
      AllowBigSelection=   0   'False
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin VB.TextBox txtProvCod 
      Height          =   315
      Left            =   195
      MaxLength       =   20
      TabIndex        =   48
      Tag             =   "txtcodigo"
      Top             =   1020
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6045
      TabIndex        =   50
      Top             =   4530
      Width           =   1905
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
      Left            =   45
      TabIndex        =   49
      Top             =   4590
      Width           =   1785
   End
End
Attribute VB_Name = "frmPlaOCIngBien"
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
Dim sObjCod As String, sObjDesc As String, sObjUnid
Dim sCtaCod As String, sCtaDesc As String
Dim sProvCod As String
Dim nTasaIGV As Currency, nVariaIGV As Currency
Dim aCtaCambio(1, 2) As String
Dim lNewProv  As Boolean
Dim sDocDesc  As String
Dim sCtaProvis As String
Dim lnMesSdoAnual As Integer
Dim lnColorBien As Double
Dim lnColorServ As Double

Dim oMov As DMov

Public Sub Inicio(LlenaObj As Boolean, OrdenCompra As Boolean, Optional ProvCod As String, Optional pnMesSdoAnual As Integer = 0)
lLlenaObj = LlenaObj
lbBienes = OrdenCompra
sProvCod = ProvCod
lnMesSdoAnual = pnMesSdoAnual
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
fgDetalle.TextMatrix(0, 4) = "Solicitado"
fgDetalle.TextMatrix(0, 5) = "P.Unitario"
fgDetalle.TextMatrix(0, 6) = "Saldo"
fgDetalle.TextMatrix(0, 7) = "Sub Total"
fgDetalle.TextMatrix(0, 8) = "Cta.Cont."
fgDetalle.TextMatrix(0, 9) = "Saldo Mes"
fgDetalle.TextMatrix(0, 11) = "Saldo Año"
fgDetalle.TextMatrix(0, 12) = "Cta Contable"

fgDetalle.ColWidth(0) = 350
fgDetalle.ColWidth(1) = 1200
fgDetalle.ColWidth(2) = 2500
fgDetalle.ColWidth(3) = 700
fgDetalle.ColWidth(4) = 1000
fgDetalle.ColWidth(5) = 1000
fgDetalle.ColWidth(6) = 0
fgDetalle.ColWidth(7) = 1100
fgDetalle.ColWidth(8) = 0
fgDetalle.ColWidth(9) = 1200
fgDetalle.ColWidth(10) = 0
fgDetalle.ColWidth(11) = 1200
fgDetalle.ColWidth(12) = 2000

fgDetalle.ColAlignment(1) = 1
fgDetalle.ColAlignment(12) = 1
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
  Dim oPrevio As clsPrevio
  Set oPrevio = New clsPrevio
  
  sTexto = "N O T A   D E   I N G R E S O    Nro. " & gcDocNro
  rtxtAsiento.Text = ImpreCabAsiento(sTexto, gdFecSis, gcEmpresaLogo, gcOpeCod, txtMovNro, txtMovDesc, False)
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
    sTexto = sTexto & Left(fgDetalle.TextMatrix(n, 1) + Space(18), 18)
    sTexto = sTexto & "  " & Left(fgDetalle.TextMatrix(n, 2) + Space(63), 63)
    sTexto = sTexto & " " & Left(fgDetalle.TextMatrix(n, 3) + Space(10), 6)
    sTexto = sTexto & " " & Right(Space(10) & fgDetalle.TextMatrix(n, 4), 10)
    sTexto = sTexto & " " & Right(Space(12) & fgDetalle.TextMatrix(n, 5), 12)
    sTexto = sTexto & " " & Right(Space(12) & fgDetalle.TextMatrix(n, 7), 12) & oImpresora.gPrnSaltoLinea
Next
sTexto = sTexto & COFF
rtxtAsiento.Text = rtxtAsiento.Text & ImpreDetLog(sTexto, "Unidad Cantidad   P.Unitario   SubTotal")
rtxtAsiento.Text = rtxtAsiento.Text & oImpresora.gPrnSaltoLinea

rtxtAsiento.Text = rtxtAsiento.Text & BON
rtxtAsiento.Text = rtxtAsiento.Text & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
rtxtAsiento.Text = rtxtAsiento.Text & "      _____________________                        _____________________" & oImpresora.gPrnSaltoLinea
rtxtAsiento.Text = rtxtAsiento.Text & "         Vo Bo ALMACEN                                Vo Bo LOGISTICA   " & BOFF & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoPagina

    rtxtAsiento.Text = ImpreCarEsp(rtxtAsiento.Text)
    oPrevio.Show rtxtAsiento.Text, "Documento: Nota de Ingreso", False, gnLinPage
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
      nTasaImp = nTasaImp + nVal(fgImp.TextMatrix(m, 2))
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
            fgDetalle.TextMatrix(n, m + 10) = Format(nImp, gcFormView)
         Else
            nImp = nVal(fgDetalle.TextMatrix(n, m + 10))
         End If
         nITot = nITot + nImp
      Else
         If lCalcImpuestos Then fgDetalle.TextMatrix(n, m + 10) = ""
      End If
   Next
   If fgImp.TextMatrix(m, 7) = "0" Then
      nTotImp = nTotImp + nITot * IIf(fgImp.TextMatrix(m, 5) = "D", 1, -1)
   End If
   fgImp.TextMatrix(m, 6) = Format(nITot, gcFormView)
   nTot = nTot + nITot * IIf(fgImp.TextMatrix(m, 5) = "D", 1, -1)
Next
txtTotal = Format(nSTot + nTotImp, gcFormView)
End Sub

Private Sub cboDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtFacSerie.SetFocus
End If
End Sub

Private Sub cmdExaCab_Click()
'Dim sSqlProv As String
'Dim rsProv As ADODB.Recordset
'Set rsProv = New ADODB.Recordset
'Dim oCon As DConecta
'Set oCon = New DConecta
'oCon.AbreConexion
'
'frmBuscaCli.Inicia frmLogIngBien, True
'lNewProv = False
'If Len(txtProvCod.Text) <> 0 Then
'   If txtProvRuc.Text = "" Then
'      txtProvRuc = "00000000"
'   Else
'      sSqlProv = "Select cCodPers from  proveedor where cCodPers = '" & txtProvCod & "'"
'      Set rsProv = oCon.CargaRecordSet(sSqlProv)
'      If RSVacio(rsProv) Then
'         lNewProv = True
'      End If
'      If rsProv.State = adStateOpen Then rsProv.Close: Set rsProv = Nothing
'   End If
'   txtMovDesc.SetFocus
'End If
End Sub

Private Sub cmdRechazar_Click()
    Dim sMovNro As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oMov As DMov
    Set oMov = New DMov
    
    On Error GoTo ErrAceptar
    
        If MsgBox(" ¿ Seguro de Rechazar Orden ? ", vbOKCancel, "Aviso de Confirmación") = vbCancel Then
            Exit Sub
        End If
        
        If Me.txtComentario.Text = "" Then
            MsgBox "Debe ingresar un comentario por el rechazo", vbInformation
            Me.txtComentario.SetFocus
            Exit Sub
        End If
        
        oCon.AbreConexion
        
        If lTransActiva Then
            oCon.RollbackTrans
           lTransActiva = False
        End If
        
        oCon.BeginTrans
          lTransActiva = True
          oMov.ActualizaMov gcMovNro, , 14
          sMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
          oMov.InsertaMov sMovNro, gsLogRechazoOCOS, Me.txtComentario.Text, gMovEstContabNoContable
          oMov.InsertaMovRef oMov.GetnMovNro(sMovNro), gcMovNro
        oCon.CommitTrans
        
        lTransActiva = False
           OK = True
           MsgBox IIf(lbBienes, "Orden de Compra", "Orden de Servicio") & " fue Rechazada", vbInformation, "¡Aviso!"
           Unload Me
        Exit Sub
ErrAceptar:
         If lTransActiva Then
             oMov.RollbackTrans
         End If
         lTransActiva = False
         MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub fgImp_Click()
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
    Unload Me
End Sub

Private Sub fgDetalle_KeyUp(KeyCode As Integer, Shift As Integer)
If fgDetalle.Col > 10 Then
   If fgImp.TextMatrix(fgDetalle.Col - 10, 0) = "." Then
      KeyUp_Flex fgDetalle, KeyCode, Shift
      CalculaTotal False
   End If
End If
End Sub

Private Sub fgDetalle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   Select Case fgDetalle.Col
          Case 1:
          Case 4: 'mnuNoAtender.Enabled = False
                  'mnuAtender.Enabled = True
                  'PopupMenu mnuLog
          Case 5: 'mnuAtender.Enabled = False
                  'mnuNoAtender.Enabled = True
                  'PopupMenu mnuLog
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
    Dim sCtaCnt As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Set oMov = New DMov
    Dim oDoc As DOperacion
    Set oDoc = New DOperacion
    Dim ofun As NContFunciones
    Set ofun = New NContFunciones
    lSalir = False
    oCon.AbreConexion
    TabDocRef.TabVisible(1) = False
    TabDocRef.TabVisible(2) = False
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
    
    Set rs = oDoc.CargaOpeCta(gcOpeCod, "H")
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
       If Mid(gcOpeCod, 3, 1) = 2 Then
         Frame2.BackColor = &H80FF80
       End If
    Else
       fgDetalle.BackColor = lnColorServ
       fgDetalle.BackColorBkg = lnColorServ
       If Mid(gcOpeCod, 3, 1) = 2 Then
         Frame2.BackColor = &H80FF80
       End If
    End If
    If sProvCod <> "" Then
       txtProvCod = sProvCod
       sSQL = " Select cPersNombre cNomPers, ISNULL(cPersIDNro,'') cProvRuc from Persona PE  Left Join PersID PID On PE.cPersCod = PID.cPersCod And Pid.cPersIDTpo = '2' Where PE.cPersCod = '" & txtProvCod & "' "
       Set rs = oCon.CargaRecordSet(sSQL)
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
       'mc.nMovNro, mc.nMovItem, mo.nMovObjOrden, mc.cCtaContCod, mo.cObjetoCod, mo.cObjetoDesc, mcd.cDescrip, moc.nMovCant, mc.nMovImporte
       If Mid(gcOpeCod, 3, 1) = "2" Then  'Identificación de Tipo de Moneda
          sSQL = "SELECT mc.nMovNro, mc.nMovItem, mc.cCtaContCod, mo.nMovbsOrden, mo.cBSCod, mo.cBSDescripcion, mcd.cDescrip, moc.nMovCant, mc.nMovImporte, mo.cConsDescripcion " _
               & "FROM MovCta mc Left Join MovME MME On mc.nMovNro = MME.nMovNro and mc.nMovItem = MME.nMovItem " _
               & "          LEFT JOIN (SELECT mo.nMovNro, mo.nMovItem, mo.nMovBsOrden, mo.cBSCod, cBSDescripcion, co.cConsDescripcion FROM MovBS mo Inner JOIN BienesServicios o  ON o.cBSCod = mo.cBSCod Inner Join Constante CO On CO.nConsValor = o.nBSUnidad And CO.nConsCod = '1019' WHERE nMovNro = '" & gcMovNro & "'" _
               & "                    ) mo on mo.nMovNro = mc.nMovNro and mo.nMovItem = mc.nMovItem " _
               & "          LEFT JOIN MovCant moc ON moc.nMovNro = mc.nMovNro and moc.nMovItem = mc.nMovItem " _
               & "          LEFT JOIN MovCotizacDet mcd ON mcd.nMovNro = mc.nMovNro and mcd.nMovItem = mc.nMovItem " _
               & "WHERE mc.nMovNro = '" & gcMovNro & "' and mc.nMovImporte <> 0 And mc.cCtaContCod Not Like '25%'"
       Else
          sSQL = "SELECT mc.nMovNro, mc.nMovItem, mc.cCtaContCod, mo.nMovBsOrden, mo.cBSCod, mo.cBSDescripcion, mcd.cDescrip, moc.nMovCant, mc.nMovImporte, mo.cConsDescripcion " _
               & "FROM MovCta mc " _
               & "          LEFT JOIN (SELECT mo.nMovNro, mo.nMovItem, mo.nMovBsOrden, mo.cBSCod, cBSDescripcion, co.cConsDescripcion FROM MovBS mo Inner JOIN BienesServicios o  ON o.cBSCod = mo.cBSCod Inner Join Constante CO On CO.nConsValor = o.nBSUnidad And CO.nConsCod = '1019' WHERE nMovNro = '" & gcMovNro & "'" _
               & "                    ) mo on mo.nMovNro = mc.nMovNro and mo.nMovItem = mc.nMovItem " _
               & "          LEFT JOIN MovCant moc ON moc.nMovNro = mc.nMovNro and moc.nMovItem = mc.nMovItem " _
               & "          LEFT JOIN MovCotizacDet mcd ON mcd.nMovNro = mc.nMovNro and mcd.nMovItem = mc.nMovItem " _
               & "WHERE mc.nMovNro = '" & gcMovNro & "' and mc.nMovImporte <> 0 And mc.cCtaContCod Not Like '25%'"
       End If
            
       Set rs = oCon.CargaRecordSet(sSQL)
       n = 0
       Do While Not rs.EOF
          n = n + 1
          If n > fgDetalle.Rows - 1 Then
            AdicionaRow fgDetalle
          End If
          fgDetalle.TextMatrix(n, 0) = n
          fgDetalle.TextMatrix(n, 7) = Format(rs!nMovImporte, gcFormView)
          fgDetalle.TextMatrix(n, 8) = rs!cCtaContCod
          '**************************************************************************************
          Dim nMovMesAcum As Currency, nMovMes As Currency
          Dim nPreMesAcum As Currency, nPreMes As Currency
          'Calcula Presupuesto pendiente de la Cta Contable
          sCtaCnt = CargaCtaPre(gdFecSis, rs!cCtaContCod)
          nPreMesAcum = CalculaPreCta(gdFecSis, sCtaCnt, True)
          nPreMes = CalculaPreCta(gdFecSis, sCtaCnt, False)
          If Not sCtaCnt = "" Then
            'Prepara para CADA CTACNT su MOV
            Dim sCC As String
            Dim nPos As Variant
            sCC = sCtaCnt
            
            nMovMesAcum = 0
            nMovMes = 0
            Do While True
                nPos = InStr(sCC, ",")
                If nPos = 0 Then
                    nMovMesAcum = nMovMesAcum + CalculaMovCta(gdFecSis, sCC)
                    nMovMes = nMovMes + CalculaMovCta(gdFecSis, sCC, False)
                    Exit Do
                Else
                    nMovMesAcum = nMovMesAcum + CalculaMovCta(gdFecSis, Left(sCC, nPos - 1))
                    nMovMes = nMovMes + CalculaMovCta(gdFecSis, Left(sCC, nPos - 1), False)
                    sCC = Mid(sCC, nPos + 1)
                End If
            Loop
          Else
             nMovMesAcum = 0
             nMovMes = 0
          End If
          fgDetalle.TextMatrix(n, 9) = nPreMes - nMovMes
          fgDetalle.TextMatrix(n, 11) = nPreMesAcum - nMovMesAcum
          fgDetalle.TextMatrix(n, 12) = Trim(rs!cCtaContCod) & "(" & sCtaCnt & ")"
          '**************************************************************************************
          If Not IsNull(rs!cDescrip) Then
             fgDetalle.TextMatrix(n, 2) = rs!cDescrip
          Else
             If Not IsNull(rs!cBSDescripcion) Then
                fgDetalle.TextMatrix(n, 2) = rs!cBSDescripcion
             End If
          End If
          If Not IsNull(rs!nMovCant) Then
             fgDetalle.TextMatrix(n, 1) = rs!cBSCod
             fgDetalle.TextMatrix(n, 4) = rs!nMovCant
             fgDetalle.TextMatrix(n, 3) = rs!cConsDescripcion
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
                   fgObj.TextMatrix(fgObj.Row, 1) = rs!nMovBsOrden
                   fgObj.TextMatrix(fgObj.Row, 2) = rs!cBSCod
                   fgObj.TextMatrix(fgObj.Row, 3) = rs!cBSDescripcion
                End If
                rs.MoveNext
                If rs.EOF Then
                   Exit Do
                End If
             Loop
          End If
       Loop
       sSQL = "SELECT dMovPlazo, cMovLugarEntrega FROM MovCotizac WHERE nMovNro = '" & gcMovNro & "'"
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
       Set rs = oDoc.CargaOpeDoc(gcOpeCod, "3")
       If RSVacio(rs) Then
          MsgBox "No se definió Documento NOTA DE INGRESO en Operación. Por favor Consultar con Sistemas...!", vbInformation, "¡Aviso!"
          lSalir = True
          Exit Sub
       Else
          gcDocTpo = rs!nDocTpo
          gcDocNro = ofun.GeneraDocNro(CInt(gcDocTpo), IIf(Mid(gcOpeCod, 3, 1) = "1", gMonedaNacional, gMonedaExtranjera))
          sDocDesc = rs!cDocDesc
       End If
       
       'GUIA DE REMISION
       Set rs = oDoc.CargaOpeDoc(gcOpeCod, "3")
       If RSVacio(rs) Then
          MsgBox "No se definió Documento GUIA DE REMISION en Operación. Por favor Consultar con Sistemas...!", vbInformation, "¡Aviso!"
          lSalir = True
          Exit Sub
       Else
          txtGRNro.Tag = rs!nDocTpo
       End If
       'lblDoc.Caption = "   " & UCase(sDocDesc) & "   Nº " & gcDocNro
    Else
       TabDocRef.TabCaption(0) = "Orden de Servicio"
       'lblDoc.Visible = False
    End If
    fgObj.BackColor = lnColorServ
    'Tipos de Comprobantes de Pago
    Set rs = oDoc.CargaOpeDoc(gcOpeCod, "22", "3")
    Do While Not rs.EOF
       cboDoc.AddItem (rs!cDocDesc & Space(100) & rs!cDocTpo)
       rs.MoveNext
    Loop
    If cboDoc.ListCount = 1 Then
       cboDoc.ListIndex = 0
    End If
    RSClose rs
    Call CalculaTotal
    
    If gcOpeCod = 501221 Or gcOpeCod = 502221 Or gcOpeCod = 501222 Or gcOpeCod = 502222 Then
       Me.Caption = "PRESUPUESTO"
    End If
End Sub

Private Function CalculaMovCta(ByVal pdFecha As Date, ByVal psCtaCont As String, _
                               Optional pbMovMesesAcum As Boolean = True) As Currency
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    Dim nMovimi As Currency
    Dim pbCtaDeu As Boolean
    Dim sCtaCont As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    'Para Centralizado ************
    'oCon.AbreConexion
    
    'Para Distribuido  ************
    'oCon.AbreConexionRemota "07", , , "04"
    oCon.AbreConexion
    If Len(psCtaCont) <= 2 Then 'enero
       sCtaCont = Trim(Left(psCtaCont, 2))
    Else
       sCtaCont = Trim(Left(psCtaCont, 2) & "_" & Mid(psCtaCont, 4))
    End If
    
'-------------Esto He Comentado
    'Verifica si es Cta Deudora
    'Distribuido
    'tmpSql = "SELECT cCodTab FROM DbComunes.dbo.TablaCod " & _
    '       " WHERE cCodTab LIKE 'C0__' AND cAbrev = 'D' " & _
    '       "       AND cValor = SubString('" & sCtaCont & "',1,Len(cValor))"

   'Centralizado
    tmpSql = "SELECT cCtaContCod FROM CtaContclase WHERE cCtaCaracter = 'D' AND cCtaContCod  = SubString('" & sCtaCont & "',1,Len(cCtaContCod ))"
    
    Set tmpReg = oCon.CargaRecordSet(tmpSql)
    If (tmpReg.BOF Or tmpReg.EOF) Then
        pbCtaDeu = False
    Else
        pbCtaDeu = True
    End If
    RSClose tmpReg
    
'    If pbMovMesesAcum Then
'        If pbCtaDeu Then
'            'CUENTA DEUDORA
'            tmpSql = "SELECT IsNull(Sum(nDebe - nHaber),0) nMov " & _
'                " FROM DBAdmin..BalanceEstad " & _
'                " WHERE cBalanceCate = '1' AND cBalanceTipo = '0' " & _
'                "   AND cCtaContCod LIKE '" & sCtaCont & "' " & _
'                "   AND cBalanceAnio = '" & Year(pdFecha) & "' " & _
'                "   AND Convert(int,cBalanceMes) < " & Month(pdFecha) & " "
'        Else
'            'CUENTA ACREEDORA
'            tmpSql = "SELECT IsNull(Sum(nHaber - nDebe),0) nMov " & _
'                " FROM DBAdmin..BalanceEstad " & _
'                " WHERE cBalanceCate = '1' AND cBalanceTipo = '0' " & _
'                "   AND cCtaContCod LIKE '" & sCtaCont & "' " & _
'                "   AND cBalanceAnio = '" & Year(pdFecha) & "'" & _
'                "   AND Convert(int,cBalanceMes) < " & Month(pdFecha) & " "
'        End If
'
'        Set tmpReg = oCon.CargaRecordSet(tmpSql)
'
'        If (tmpReg.BOF Or tmpReg.EOF) Then
'            nMovimi = 0
'        Else
'            nMovimi = tmpReg!nMov
'        End If
'        RSClose tmpReg
'    Else
        'Mes actual

'----------------------ESTO DESCOMENTE-----------------
        'Para Base Centralizada *****************
        tmpSql = " SELECT IsNull(Sum(mc.nMovImporte),0) nMov" _
               & " FROM MovCta MC" _
               & " JOIN MOV M ON mc.nMovNro = m.nMovNro" _
              & " WHERE m.cMovNro LIKE '" & Left(Format(pdFecha, gsFormatoMovFecha), 6) & "%'" _
               & "   AND mc.cCtaContCod like '" & sCtaCont & "%'" _
               & "   AND m.nMovFlag <> '1'" _
               & "   AND (m.nMovEstado='16' OR (m.nMovEstado='0' AND Not" _
               & " Exists( Select Distinct m2.nMovEstado From mov m2 join movref mr On m2.nMovNro = mr.nMovNroref And mr.nMovNro = mc.nMovNro Where m2.nMovFlag <> '3' And m2.nMovEstado = '16')))"
               
'------------------------COMENTE ESTO ----------------------------
        'Para Base Distribuida con Centralizada *******************
        'tmpSql = "SELECT SUM(nMov) nMov FROM ( SELECT IsNull(Sum(mc.nMovImporte),0) nMov" & _
        '   " FROM  DBAdmin..MovCta MC JOIN DBADmin..MOV M ON mc.cMovNro = m.cMovNro " & _
        '    " WHERE LEFT(m.cMovNro,4) = '" & Year(pdFecha) & "' and SubString(m.cMovNro,5,2) " & IIf(pbMovMesesAcum, "<=", "=") & " '" & Format(Month(pdFecha), "00") & "' " & _
        '    "   AND mc.cCtaContCod like '" & sCtaCont & "%'  AND SubString(mc.cCtaContCod,3,1) IN ('1','2') " & _
        '    "   AND m.cMovFlag <> 'X' " & _
        '    "   AND (m.cMovEstado='9' OR (m.cMovEstado='0' AND Not Exists(SELECT Distinct m2.nMovEstado " & _
        '    "                                        FROM   Mov m2 JOIN DBADmin..movref mr On m2.cmovnro = mr.cmovnroref " & _
        '    "                                           And mr.cMovNro = mc.cMovNro " & _
        '    "                                        WHERE m2.nMovFlag <> '" & gMovFlagEliminado & "' And m2.nMovEstado = '" & gMovEstPresupAceptado & "'))) UNION ALL " & _
        '    "   SELECT IsNull(Sum(mc.nMovImporte),0) nMov " & _
        '    "   FROM MovCta MC" & _
        '    "   JOIN MOV M ON mc.nMovNro = m.nMovNro" & _
        '    "   WHERE LEFT(m.cMovNro,4) = '" & Year(pdFecha) & "' and SubString(m.cMovNro,5,2) " & IIf(pbMovMesesAcum, "<=", "=") & " '" & Format(Month(pdFecha), "00") & "' " & _
        '    "    AND mc.cCtaContCod like '" & sCtaCont & "%' AND SubString(mc.cCtaContCod,3,1) IN ('1','2') " & _
        '    "    AND m.nMovFlag <> '" & gMovFlagEliminado & "' and nMovEstado = '" & gMovEstPresupAceptado & "' ) Movimiento "

        Set tmpReg = oCon.CargaRecordSet(tmpSql)
        If (tmpReg.BOF Or tmpReg.EOF) Then
            nMovimi = 0
        Else
            If pbCtaDeu Then
                nMovimi = tmpReg!nmov
            Else
                nMovimi = (tmpReg!nmov * -1)
            End If
        End If
        RSClose tmpReg
'    End If
    CalculaMovCta = nMovimi
End Function

Private Function CalculaPreCta(ByVal pdFecha As Date, ByVal psCtaCont As String, _
Optional pbMovMesesAcum As Boolean = True) As Currency
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    Dim nMovimi As Currency
    Dim nPaso As Integer
    Dim nLongitud As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    'Prepara para cada ctacnt su Like
    Dim sCtaCont As String
    Dim nPos As Variant
    Dim sCad As String
    sCtaCont = psCtaCont
    sCad = ""
    Do While True
        nPos = InStr(sCtaCont, ",")
        If nPos = 0 Then
            sCad = sCad & " AND rc.cCtaContCod LIKE '" & sCtaCont & "' "
            Exit Do
        Else
            sCad = sCad & " AND rc.cCtaContCod LIKE '" & Left(sCtaCont, nPos - 1) & "' "
            Exit Do
        End If
    Loop

    If pbMovMesesAcum Then
        'Carga el Presupuesto Acumulado
        Dim lnMes As Integer
        lnMes = Int(Month(pdFecha) / lnMesSdoAnual) + IIf(Month(pdFecha) Mod lnMesSdoAnual = 0, 0, 1)
        lnMes = lnMes * lnMesSdoAnual
        tmpSql = "SELECT IsNull(Sum(p.nPresuRubMesMonIni),0) nMov  FROM PresuRubroMes P JOIN PresuRubroCta RC ON p.nPresuCod = rc.nPresuCod and p.cPresuRubCod Like rc.cPresuRubCod + '%' AND p.nPresuAnio = rc.nPresuAnio WHERE p.nPresuAnio =  " & Year(pdFecha) & " AND P.nPresuMes <= " & lnMes & sCad
        Set tmpReg = oCon.CargaRecordSet(tmpSql)
        If (tmpReg.BOF Or tmpReg.EOF) Then
            nMovimi = 0
        Else
            nMovimi = Abs(tmpReg!nmov)
        End If
    Else
        'Presupuesto del Mes actual
        tmpSql = "SELECT IsNull(Sum(p.nPresuRubMesMonIni),0) nMov  FROM PresuRubroMes P JOIN PresuRubroCta RC ON p.nPresuCod = rc.nPresuCod and p.cPresuRubCod Like rc.cPresuRubCod + '%' AND p.nPresuAnio = rc.nPresuAnio WHERE p.nPresuAnio = " & Year(pdFecha) & " AND P.nPresuMes = " & Month(pdFecha) & sCad
        Set tmpReg = oCon.CargaRecordSet(tmpSql)
        
        If (tmpReg.BOF Or tmpReg.EOF) Then
            nMovimi = 0
        Else
            nMovimi = Abs(tmpReg!nmov)
        End If
    End If
    RSClose tmpReg
    
    CalculaPreCta = nMovimi
End Function

'Carga la cuenta a buscar en Presupuesto y en Moviminetos
Private Function CargaCtaPre(ByVal pdFecha As Date, ByVal psCtaCont As String) As String
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    Dim sCtaCont As String
    Dim nPaso As Integer
    Dim nLongitud As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
    nPaso = 0
    nLongitud = 0
    Do While True
        nLongitud = Len(psCtaCont) - 2 - (nPaso * 2)
        If nLongitud <= 0 Then
            If nLongitud < 0 Then
                CargaCtaPre = ""
                Exit Function
            End If
            sCtaCont = Trim(Left(psCtaCont, 2)) & ""
        Else
            sCtaCont = Trim(Left(psCtaCont, 2) & "_" & Mid(psCtaCont, 4, Len(psCtaCont) - 3 - (nPaso * 2))) & ""
        End If
        nPaso = nPaso + 1
        
        tmpSql = " SELECT distinct(rc.cCtaContCod)" _
               & " FROM PresuRubroMes P" _
               & " JOIN PresuRubroCta RC" _
               & " ON P.nPresuCod = RC.nPresucod and P.cPresuRubCod Like RC.cPresuRubCod + '%'" _
               & " AND P.nPresuAnio = RC.nPresuAnio" _
               & " WHERE p.nPresuAnio = " & Year(pdFecha) & " AND P.nPresuMes < " & Month(pdFecha) & " AND rc.cCtaContCod LIKE '" & sCtaCont & "'"
        Set tmpReg = oCon.CargaRecordSet(tmpSql)
        If tmpReg.RecordCount = 0 Then
        Else
            sCtaCont = ""
            'Salir del Bucle
            Do While Not tmpReg.EOF
                If Len(tmpReg!cCtaContCod) <= 2 Then
                    sCtaCont = sCtaCont & "," & Left(tmpReg!cCtaContCod, 2)
                Else
                    sCtaCont = sCtaCont & "," & Left(tmpReg!cCtaContCod, 2) & "_" & Mid(tmpReg!cCtaContCod, 4)
                End If
                tmpReg.MoveNext
            Loop
            Exit Do
        End If
        tmpReg.Close
    Loop
    tmpReg.Close
    Set tmpReg = Nothing
    
    CargaCtaPre = Mid(sCtaCont, 2)
End Function



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

Private Sub txtComentario_GotFocus()
    Me.txtComentario.SelStart = 0
    Me.txtComentario.SelLength = 500
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdRechazar.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
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
'If ValidaFecha(txtFecha.Text) <> "" Then
'   MsgBox " Fecha no Válida ", vbInformation, "Aviso"
'   Cancel = True
'Else
'   txtMovNro = GeneraMovNro(, , txtFecha)
'End If
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
'KeyAscii = intfNumEnt(KeyAscii)
'If KeyAscii = 13 Then
'   txtGRNro = Right(String(8, "0") & txtGRNro, 8)
'   txtGRFecha.SetFocus
'End If
End Sub
Private Sub txtFacNro_KeyPress(KeyAscii As Integer)
'KeyAscii = intfNumEnt(KeyAscii)
If KeyAscii = 13 Then
   txtFacNro = Right(String(8, "0") & txtFacNro, 8)
   txtFacFecha.SetFocus
End If
End Sub

Private Sub txtGRSerie_KeyPress(KeyAscii As Integer)
'KeyAscii = intfNumEnt(KeyAscii)
If KeyAscii = 13 Then
   txtGRSerie = Right(String(3, "0") & txtGRSerie, 3)
   txtGRNro.SetFocus
End If
End Sub
Private Sub txtFacSerie_KeyPress(KeyAscii As Integer)
'KeyAscii = intfNumEnt(KeyAscii)
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

Private Sub cmdAceptar_Click()
    Dim n As Integer 'Contador
    Dim nItem As Integer, nCol  As Integer
    Dim sTexto As String, lOk As Boolean
    Dim sMovNro As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    On Error GoTo ErrAceptar
    
    If MsgBox(" ¿ Seguro de grabar Operación ? ", vbOKCancel, "Aviso de Confirmación") = vbCancel Then
        Exit Sub
    End If
    
    If Me.txtComentario.Text = "" Then
        MsgBox "Debe ingresar un comentario por la Aprobación.", vbInformation
        Me.txtComentario.SetFocus
        Exit Sub
    End If
    
    oCon.AbreConexion
    ' Iniciamos transaccion
    If lTransActiva Then
       lTransActiva = False
    End If
    oCon.BeginTrans
      lTransActiva = True
      oMov.ActualizaMov gcMovNro, , 16
      sMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
      oMov.InsertaMov sMovNro, gsLogApruebaOCOS, Me.txtComentario.Text, gMovEstContabNoContable
      oMov.InsertaMovRef oMov.GetnMovNro(sMovNro), gcMovNro
    oCon.CommitTrans
    lTransActiva = False
       OK = True
    '   If lbBienes Then
    '      ImprimeNotaIngreso
    '   End If
       MsgBox IIf(lbBienes, "Orden de Compra", "Orden de Servicio") & " Aceptada", vbInformation, "¡Aviso!"
       Unload Me
    Exit Sub
ErrAceptar:
    If lTransActiva Then
        oCon.RollbackTrans
    End If
    lTransActiva = False
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Public Property Get lOk() As Boolean
lOk = OK
End Property

Public Property Let lOk(ByVal vNewValue As Boolean)
OK = vNewValue
End Property

