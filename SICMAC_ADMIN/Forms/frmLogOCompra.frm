VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogOCompra 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   930
   ClientTop       =   1260
   ClientWidth     =   11130
   Icon            =   "frmLogOCompra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   11130
   Visible         =   0   'False
   Begin VB.CommandButton cmdDocumentoPDF 
      Caption         =   "Imp. &PDF"
      Height          =   315
      Left            =   7560
      TabIndex        =   80
      Top             =   6120
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Frame fraComprobante 
      Caption         =   "Comprobante"
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
      Height          =   1395
      Left            =   11280
      TabIndex        =   69
      Top             =   1080
      Visible         =   0   'False
      Width           =   3480
      Begin VB.CommandButton cmdDefinir 
         Caption         =   "&Definir"
         Height          =   375
         Left            =   2160
         TabIndex        =   75
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cboDoc 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   240
         Width           =   2460
      End
      Begin VB.TextBox txtFacSerie 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         MaxLength       =   4
         TabIndex        =   71
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtFacNro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         MaxLength       =   12
         TabIndex        =   70
         Top             =   600
         Width           =   1590
      End
      Begin MSComCtl2.DTPicker txtEmision 
         Height          =   315
         Left            =   840
         TabIndex        =   76
         Top             =   960
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   238354433
         CurrentDate     =   37156
      End
      Begin VB.Label lblEmision 
         AutoSize        =   -1  'True
         Caption         =   "Emisión:"
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label lblTipoComp 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   360
         TabIndex        =   74
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblNComp 
         AutoSize        =   -1  'True
         Caption         =   "Nº:"
         Height          =   195
         Left            =   480
         TabIndex        =   72
         Top             =   600
         Width           =   225
      End
   End
   Begin VB.CommandButton cmdExaminarDes 
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   4320
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox textObjDes 
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
      Left            =   5880
      TabIndex        =   67
      Top             =   4320
      Visible         =   0   'False
      Width           =   1510
   End
   Begin VB.CommandButton Command1 
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
      Left            =   7110
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   4335
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.CheckBox chkRenta4 
      Caption         =   "Renta 4ta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   64
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkIGV 
      Caption         =   "IGV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   63
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame fraCot 
      Height          =   660
      Left            =   7530
      TabIndex        =   39
      Top             =   1650
      Width           =   3465
      Begin VB.TextBox txtCotNro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         MaxLength       =   15
         TabIndex        =   5
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label2 
         Caption         =   "Cotización Nro."
         Height          =   225
         Left            =   150
         TabIndex        =   40
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.TextBox txtProvRuc 
      Height          =   315
      Left            =   6585
      MaxLength       =   20
      TabIndex        =   52
      Tag             =   "txtcodigo"
      Top             =   5595
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
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
      Height          =   705
      Left            =   3840
      TabIndex        =   50
      Top             =   15
      Width           =   1740
      Begin VB.ComboBox cmbMoneda 
         Enabled         =   0   'False
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   285
         Width           =   1500
      End
   End
   Begin MSMask.MaskEdBox txtFecPlazo 
      Height          =   315
      Left            =   8280
      TabIndex        =   46
      Top             =   4305
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BackColor       =   15794175
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.CheckBox chkProy 
      Caption         =   "Afecta a Proyectos"
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
      Left            =   11400
      TabIndex        =   45
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame fraProy 
      Enabled         =   0   'False
      Height          =   705
      Left            =   11280
      TabIndex        =   42
      Top             =   240
      Width           =   3465
      Begin VB.CommandButton cmdProy 
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
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   240
         Width           =   270
      End
      Begin VB.TextBox txtProy 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   240
         Width           =   3120
      End
   End
   Begin VB.TextBox txtConcepto 
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
      Height          =   300
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   4275
      Visible         =   0   'False
      Width           =   1800
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
      TabIndex        =   11
      Top             =   4260
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txtDescServ 
      Height          =   315
      Left            =   5310
      TabIndex        =   29
      Top             =   6960
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdServicio 
      Caption         =   "Ser&vicio"
      Height          =   330
      Left            =   7560
      TabIndex        =   16
      Top             =   6120
      Width           =   1110
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
      Left            =   4350
      TabIndex        =   13
      Top             =   4815
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
      Left            =   1800
      TabIndex        =   10
      Top             =   4245
      Visible         =   0   'False
      Width           =   1510
   End
   Begin VB.TextBox txtTot 
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
      Height          =   285
      Left            =   8730
      TabIndex        =   35
      Top             =   5670
      Width           =   1185
   End
   Begin VB.TextBox txtProvTele 
      Height          =   315
      Left            =   5130
      MaxLength       =   20
      TabIndex        =   34
      Tag             =   "txtTelefono"
      Top             =   6960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtProvDir 
      Height          =   315
      Left            =   4980
      TabIndex        =   33
      Tag             =   "txtDireccion"
      Top             =   6960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   8760
      TabIndex        =   17
      Top             =   6120
      Width           =   1110
   End
   Begin VB.PictureBox PicOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4380
      Picture         =   "frmLogOCompra.frx":030A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   28
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtProvCod 
      Height          =   315
      Left            =   4830
      MaxLength       =   20
      TabIndex        =   24
      Tag             =   "txtcodigo"
      Top             =   6960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Frame fraProv 
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
      Height          =   885
      Left            =   105
      TabIndex        =   19
      Top             =   735
      Width           =   6870
      Begin VB.TextBox txtProvNom 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Tag             =   "txtnombre"
         Top             =   210
         Width           =   5160
      End
      Begin Sicmact.TxtBuscar txtPersona 
         Height          =   315
         Left            =   135
         TabIndex        =   49
         Top             =   240
         Width           =   1380
         _extentx        =   2434
         _extenty        =   529
         appearance      =   0
         appearance      =   0
         font            =   "frmLogOCompra.frx":064C
         appearance      =   0
         tipobusqueda    =   3
         stitulo         =   ""
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1605
         TabIndex        =   53
         Top             =   540
         Width           =   4980
         Begin VB.CheckBox chkRetencion 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Retención"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   55
            Top             =   30
            Width           =   1035
         End
         Begin VB.CheckBox chkBuneCOnt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Buen Contribuyente"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2655
            TabIndex        =   54
            Top             =   30
            Width           =   1815
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fecha"
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
      Height          =   705
      Left            =   5580
      TabIndex        =   21
      Top             =   15
      Width           =   1395
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   150
         TabIndex        =   0
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   330
      Left            =   9960
      TabIndex        =   18
      Top             =   6120
      Width           =   1110
   End
   Begin RichTextLib.RichTextBox rtxtAsiento 
      Height          =   345
      Left            =   2490
      TabIndex        =   20
      Top             =   6990
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   609
      _Version        =   393217
      Enabled         =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmLogOCompra.frx":0678
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgDetalle 
      Height          =   1725
      Left            =   120
      TabIndex        =   14
      Top             =   3915
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   3043
      _Version        =   393216
      Rows            =   121
      Cols            =   13
      ForeColorSel    =   -2147483643
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483637
      AllowBigSelection=   0   'False
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgObj 
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   5925
      Width           =   7335
      _ExtentX        =   12938
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
   Begin VB.Frame fraArea 
      Caption         =   "Area Usuaria"
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
      Height          =   660
      Left            =   120
      TabIndex        =   38
      Top             =   1650
      Width           =   7365
      Begin Sicmact.TxtBuscar txtArea 
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1410
         _extentx        =   2487
         _extenty        =   556
         appearance      =   0
         appearance      =   0
         font            =   "frmLogOCompra.frx":06F8
         appearance      =   0
      End
      Begin VB.TextBox txtAgeDesc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         TabIndex        =   4
         Top             =   240
         Width           =   5640
      End
   End
   Begin VB.Frame fraForm 
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
      TabIndex        =   30
      Top             =   2310
      Width           =   10875
      Begin VB.ComboBox cboContSuministro 
         Height          =   315
         Left            =   8880
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox chkContSuministro 
         Caption         =   "De Contrato de Suministro"
         Height          =   255
         Left            =   6600
         TabIndex        =   78
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cmbTipoOC 
         Height          =   315
         ItemData        =   "frmLogOCompra.frx":0724
         Left            =   5400
         List            =   "frmLogOCompra.frx":072E
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker txtPlazo 
         Height          =   315
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   238354433
         CurrentDate     =   37156
      End
      Begin VB.ComboBox cboFormaPago 
         Height          =   315
         ItemData        =   "frmLogOCompra.frx":07A9
         Left            =   1290
         List            =   "frmLogOCompra.frx":07B9
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label lbltipooc 
         AutoSize        =   -1  'True
         Caption         =   "Tipo O/C"
         Height          =   240
         Left            =   4560
         TabIndex        =   57
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label14 
         Caption         =   "Entrega"
         Height          =   240
         Left            =   2580
         TabIndex        =   32
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label7 
         Caption         =   "Forma de Pago"
         Height          =   225
         Left            =   120
         TabIndex        =   31
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame fraMovDesc 
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
      Height          =   930
      Left            =   120
      TabIndex        =   23
      Top             =   2955
      Width           =   10875
      Begin VB.CommandButton cmdSeleccion 
         Caption         =   "Seleccion"
         Height          =   375
         Left            =   -1275
         TabIndex        =   61
         Top             =   150
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtDiasAtraso 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2025
         MaxLength       =   40
         TabIndex        =   59
         Top             =   555
         Width           =   450
      End
      Begin VB.TextBox txtMovDesc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   150
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   210
         Width           =   8715
      End
      Begin VB.Label Label3 
         Caption         =   "meses"
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
         Left            =   2520
         TabIndex        =   62
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblDiasAtraso 
         AutoSize        =   -1  'True
         Caption         =   "Periodo  de Garantia"
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
         Left            =   180
         TabIndex        =   60
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.CommandButton cmdDocumento 
      Caption         =   "&Imprimir"
      Height          =   315
      Left            =   8760
      TabIndex        =   47
      Top             =   6165
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Frame fraCambio 
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
      Height          =   885
      Left            =   7080
      TabIndex        =   25
      Top             =   735
      Visible         =   0   'False
      Width           =   3465
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
         TabIndex        =   2
         Top             =   330
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
         Left            =   2280
         TabIndex        =   3
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Fijo"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   390
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "Compra"
         Height          =   255
         Left            =   1620
         TabIndex        =   26
         Top             =   390
         Width           =   555
      End
   End
   Begin VB.Frame fraLug 
      Caption         =   "Lugar de Entrega"
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
      Left            =   600
      TabIndex        =   41
      Top             =   4200
      Width           =   4185
      Begin VB.TextBox txtLugarEntrega 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1620
         MaxLength       =   40
         TabIndex        =   6
         Top             =   240
         Width           =   2400
      End
   End
   Begin Sicmact.TxtBuscar txtCodLugarEntrega 
      Height          =   315
      Left            =   4440
      TabIndex        =   65
      Top             =   2520
      Width           =   1215
      _extentx        =   2355
      _extenty        =   556
      appearance      =   0
      appearance      =   0
      backcolor       =   15794175
      font            =   "frmLogOCompra.frx":0802
      appearance      =   0
      stitulo         =   ""
   End
   Begin VB.Label lblDocOCD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ORDEN DE COMPRA DIRECTA   Nº 00000000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   660
      Left            =   6960
      TabIndex        =   56
      Top             =   60
      Visible         =   0   'False
      Width           =   4080
   End
   Begin VB.Label Label4 
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
      Left            =   135
      TabIndex        =   37
      Top             =   5715
      Width           =   1785
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL"
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
      Left            =   7515
      TabIndex        =   36
      Top             =   5670
      Width           =   1230
   End
   Begin VB.Label lblDoc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ORDEN DE COMPRA   Nº 00000000"
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
      Height          =   660
      Left            =   135
      TabIndex        =   22
      Top             =   60
      Width           =   3720
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
   End
End
Attribute VB_Name = "frmLogOCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim lTransActiva As Boolean      ' Controla si la transaccion esta activa o no
Dim rs As ADODB.Recordset    'Rs temporal para lectura de datos
Dim lSalir As Boolean
Dim lLlenaObj  As Boolean, OK As Boolean, lOrdenCompra As Boolean
Dim sObjCod    As String, sObjDesc As String, sObjUnid As String
'ALPA 20091110*******************************************
Dim sAgeCodDes As String
'********************************************************
Dim sCtaCod    As String, sCtaDesc As String
Dim sProvCod   As String
Dim sCtaProvis As String
Dim sBS        As String
Dim lNewProv   As Boolean, lbBienes As Boolean, lbModifica As Boolean
Dim sDocDesc   As String
Dim lnColorBien As Double
Dim lnColorServ As Double
Dim lsMovNro    As String
Dim lsMovNroEdit As String 'PASI20151210
Dim lbImprime   As Boolean
Dim lsEstado    As String
Dim AceptaOK As Boolean
Dim lsCodPers As String
Dim lbRegComp  As Boolean 'WIOR 20130110
Dim fsObjetoCod As String 'EJVG20140319

Dim lsDocNroOCD As String
'ARLO 20170125******************
Dim objPista As COMManejador.Pista
'*******************************

Public Sub Inicio(LlenaObj As Boolean, pcOpeCod As String, Optional ProvCod As String = "", Optional lBienes As Boolean = True, Optional pbModifica As Boolean = False, Optional pbImprime As Boolean = False, Optional psEstado As String = "", Optional psCaption As String = "", Optional pbRegComp As Boolean = False) 'WIOR 20130110 Agrego pbRegComp
    lLlenaObj = LlenaObj
    sProvCod = ProvCod
    lsCodPers = ProvCod
    lbBienes = lBienes
    lbModifica = pbModifica
    lbImprime = pbImprime
    lsEstado = psEstado
    gcOpeCod = pcOpeCod
    lbRegComp = pbRegComp 'WIOR 20130110
    Me.Show 1
End Sub

Public Function FormaSelect(psOpeCod As String, sObj As String, nNiv As Integer, psAgeCod As String) As String
    Dim sText As String
    sText = " SELECT b.cCtaContCod cObjetoCod, b.cCtaContDesc cObjetoDesc, e.cBSCod cObjCod," _
          & " upper(e.cBSDescripcion) as cObjDesc, 1 nObjetoNiv, CO.cConsDescripcion " _
          & " FROM  CtaCont b " _
          & " Inner JOIN CtaBS  c ON Replace(c.cCtaContCod,'AG','" & psAgeCod & "') = b.cCtaContCod And cOpeCod = '" & psOpeCod & "'" _
          & " Inner JOIN BienesServicios e ON e.cBSCod like c.cObjetoCod + '%'" _
          & " Inner Join Constante CO On nBSunidad = CO.nConsValor And nConsCod = '1019'" _
          & " "
    If nNiv > 0 Then
       sText = sText & "WHERE d.nObjetoNiv = " & nNiv & " "
    End If
    FormaSelect = sText & IIf(sObj <> "", " And e.cBSCod = '" & sObj & "' ", sObj) _
                & "ORDER BY e.cBSCod"
End Function

Private Function ValidarProvee(ProvRuc As String) As Boolean
'    Dim sSqlProv As String
'    Dim rsProv As New ADODB.Recordset
'    ValidarProvee = False
'    If Len(Trim(ProvRuc)) = 0 Then
'       Exit Function
'    End If
'    sSqlProv = gcCentralPers & "spBuscaClienteDoc '" & Trim(ProvRuc) & "'"
'    RSClose rsProv
'    rsProv.Open sSqlProv, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
'    If rsProv.EOF Then
'       MsgBox " Proveedor no Encontrado ...! ", vbCritical, "Aviso"
'    Else
'       txtProvNom = rsProv!cNomPers
'       txtProvCod = rsProv!cCodPers
'       txtProvTele = rsProv!cTelPers
'       txtProvDir = rsProv!cDirPers
'       lNewProv = False
'       ValidarProvee = True
'    End If
'    RSClose rsProv
End Function


Private Sub cboFormaPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtCodLugarEntrega.SetFocus
    End If
End Sub

Private Sub chkProy_Click()
    If Not lbImprime Then
       Exit Sub
    End If
    If chkProy.value = vbChecked Then
       fraProy.Enabled = True
       cmdProy.SetFocus
    Else
       fraProy.Enabled = False
       txtProy.Text = ""
       txtProy.Tag = ""
       txtPersona.SetFocus
    End If
End Sub





Private Sub cmbTipoOC_Click()
   
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim rs As ADODB.Recordset
   Dim sSql As String
   Set rs = New ADODB.Recordset
   Dim sTipodoc As String
   Dim sPreNombre As String
   
   Dim ofun As NContFunciones
   Set ofun = New NContFunciones
   
   If lbBienes = True Then
      If Right(cmbTipoOC.Text, 1) = "D" Then
        sTipodoc = "130"
        sPreNombre = " DIRECTA"
      ElseIf Right(cmbTipoOC.Text, 1) = "P" Then
        sTipodoc = "133"
        sPreNombre = " PROCESO"
      End If
   Else
        If Right(cmbTipoOC.Text, 1) = "D" Then
        sTipodoc = "132"
        sPreNombre = " DIRECTA"
        ElseIf Right(cmbTipoOC.Text, 1) = "P" Then
        sTipodoc = "134"
        sPreNombre = " PROCESO"
        End If
   End If
   
   If Right(cmbTipoOC.Text, 1) = "D" Or Right(cmbTipoOC.Text, 1) = "P" Then
   
   lblDocOCD.Visible = True
   
   If lbModifica = True Then
       'obtener  el numero grabado
       
        sSql = " select  cDocNro from movdoc where nMovNro = " & gcMovNro & " and nDocTpo ='" & sTipodoc & "' "
        If oCon.AbreConexion = True Then
                Set rs = oCon.CargaRecordSet(sSql)
                If rs.EOF Then
                    'le genera el numero de orden mas cercano
                    'lsDocNroOCD = GeneraDocNro(CInt(sTipodoc), Mid(gcOpeCod, 3, 1), Year(gdFecSis))
                    'lblDocOCD.Caption = UCase(sDocDesc) & sPrenombre & " Nº " & lsDocNroOCD
                    lblDocOCD.Caption = ""
                    lblDocOCD.Visible = False
                    Else
                    lblDocOCD.Caption = UCase(sDocDesc) & sPreNombre & " Nº " & rs!cDocNro
                End If
        End If
   Else
       'obtener uno nuevo
        lsDocNroOCD = ofun.GeneraDocNro(CInt(sTipodoc), Mid(gcOpeCod, 3, 1), Year(gdFecSis), "OSC") 'NAGL 20191212 Agregó "OSC"
        lblDocOCD.Caption = UCase(sDocDesc) & sPreNombre & " Nº " & lsDocNroOCD
   End If
   
   
   Else
   lblDocOCD.Visible = False
End If

End Sub

'WIOR 20130131 ****************************
Private Sub cmdDefinir_Click()
If ValidarComprobante Then
    If MsgBox("Esta seguro de guardar los Datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Dim olog As DLogGeneral
        Dim rsLog As ADODB.Recordset
        Dim nTipRel As Integer
        
        Set olog = New DLogGeneral
        
        nTipRel = 0
        Select Case gcOpeCod
            Case "501210": nTipRel = 1
            Case "502210": nTipRel = 2
            Case "501211": nTipRel = 3
            Case "502211": nTipRel = 4
        End Select
        
        Set rsLog = olog.ObtenerComprobante(gcDocNro, Trim(Right(lblDocOCD.Caption, 14)), 1, gcOpeCod)
        
        If Not (rsLog.BOF And rsLog.EOF) Then
            If rsLog.RecordCount > 0 Then
                Call olog.InsActComprobante(Trim(txtFacSerie.Text) & "-" & Trim(txtFacNro.Text), Trim(Right(cboDoc.Text, 5)), txtEmision.value, gcDocNro, Trim(Right(lblDocOCD.Caption, 14)), nTipRel, gcOpeCod, 2)
            End If
        Else
            Call olog.InsActComprobante(Trim(txtFacSerie.Text) & "-" & Trim(txtFacNro.Text), Trim(Right(cboDoc.Text, 5)), txtEmision.value, gcDocNro, Trim(Right(lblDocOCD.Caption, 14)), nTipRel, gcOpeCod, 1)
        End If
        
        MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
        
        cboDoc.Enabled = False
        txtFacNro.Enabled = False
        txtFacSerie.Enabled = False
        txtEmision.Enabled = False
        cmdDefinir.Enabled = False
        OK = True
    End If
End If
End Sub

Private Function ValidarComprobante() As Boolean
Dim olog As DLogGeneral
Set olog = New DLogGeneral

If olog.OrdenProvisionada(gcMovNro) Then
    MsgBox "Orden ya fue Provisionada", vbInformation, "Aviso"
    ValidarComprobante = False
    Exit Function
End If


If Trim(cboDoc.Text) = "" Then
    MsgBox "Seleccione el tipo de documento", vbInformation, "Aviso"
    cboDoc.SetFocus
    ValidarComprobante = False
    Exit Function
End If

If Trim(txtFacNro.Text) = "" Or Trim(txtFacSerie.Text) = "" Then
    MsgBox "Ingrese el Nº Comprobante correctamente", vbInformation, "Aviso"
    txtFacSerie.SetFocus
    ValidarComprobante = False
    Exit Function
End If

ValidarComprobante = True
End Function
'WIOR *************************************
Private Sub cmdExaminarDes_Click()
    Dim oConst As DConstantes
    Dim rs As ADODB.Recordset
    Set oConst = New DConstantes
    Dim oDesc As New ClassDescObjeto
    Set rs = New ADODB.Recordset
    Set rs = oConst.GetAgencias(, , True)
    oDesc.lbUltNivel = True
    oDesc.Show rs, textObjDes
    If oDesc.lbOK Then
       sAgeCodDes = oDesc.gsSelecCod
       fgDetalle.TextMatrix(fgDetalle.row, 1) = sAgeCodDes
        'ActualizaFG fgDetalle.Row
    End If
    fgDetalle.Enabled = True
    textObjDes.Visible = False
    cmdExaminarDes.Visible = False
    fgDetalle.SetFocus
    rs.Close
    'EJVG20140319 ***
    fsObjetoCod = ""
    If fgDetalle.TextMatrix(fgDetalle.row, 11) = "B" And Len(Trim(fgDetalle.TextMatrix(fgDetalle.row, 2))) > 0 Then
        fsObjetoCod = Trim(fgDetalle.TextMatrix(fgDetalle.row, 2))
        cmdExaminar_Click
        fsObjetoCod = ""
    End If
    'END EJVG *******
End Sub

Private Sub cmdProy_Click()
'   sSQL = "SELECT cPresu cObjetoCod, cDesPre cObjetoDesc, 1 nObjetoNiv FROM PPresu WHERE cTpo = '3' "
'   Set rs = dbCmact.Execute(sSQL)
'   If rs.EOF Then
'      MsgBox "No existen Proyectos Presupuestados para Asignar Gastos", vbInformation, "¡Aviso!"
'      RSClose rs
'      Exit Sub
'   End If
'   frmDescObjeto.Inicio rs, "", 1, "Proyectos Presupuestados"
'   If frmDescObjeto.lOk Then
'      txtProy.Tag = gaObj(0, 0, 0)
'      txtProy.Text = gaObj(0, 1, 0)
'      If txtPersona.Enabled Then
'         txtPersona.SetFocus
'      Else
'         txtAgeCod.SetFocus
'      End If
'   End If
'   RSClose rs
End Sub

Private Sub cmdSeleccion_Click()
frmLogSelPaseContra.Show vbModal
lsCodPers = txtPersona.psCodigoPersona
'Asignar Cuentas Contables

End Sub

Private Sub cmdServicio_Click()
If fgDetalle.TextMatrix(fgDetalle.row, 11) = "B" Then
   If fgDetalle.TextMatrix(fgDetalle.row, 1) <> "" Then
      MsgBox "Item asignado para Ingreso de Bienes y ya tiene un Valor Asignado", vbInformation, "¡Aviso!"
      Exit Sub
   End If
   fgDetalle.TextMatrix(fgDetalle.row, 11) = "S"
   FlexBackColor fgDetalle, fgDetalle.row, lnColorServ
   cmdServicio.Caption = "&Bienes"
Else
   If fgDetalle.TextMatrix(fgDetalle.row, 1) <> "" Then
      MsgBox "Item asignado para Ingreso de Servicios y ya tiene un Valor Asignado", vbInformation, "¡Aviso!"
      Exit Sub
   End If
   fgDetalle.TextMatrix(fgDetalle.row, 11) = "B"
   FlexBackColor fgDetalle, fgDetalle.row, lnColorBien
   cmdServicio.Caption = "&Servicio"
End If
End Sub

Private Sub fgDetalle_GotFocus()
txtObj.Visible = False
cmdExaminar.Visible = False
End Sub

Private Sub fgDetalle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu mnuObj
End If
End Sub

Private Sub fgDetalle_RowColChange()
If fgDetalle.TextMatrix(fgDetalle.row, 11) = "B" Then
   cmdServicio.Caption = "&Servicio"
Else
   cmdServicio.Caption = "&Bienes"
End If
If fgDetalle.TextMatrix(fgDetalle.row, 0) <> "" Then
    RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.row, 0)
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not OK And Not lbImprime Then
    If MsgBox("¿ Desea Salir sin Grabar Documento ?", vbQuestion + vbYesNo, "¡Aviso!") = vbNo Then
       Cancel = 1
    End If
End If
End Sub



Private Sub mnuAgregar_Click()
If fgDetalle.TextMatrix(1, 0) = "" Then
   AdicionaRow fgDetalle, 1
   EnfocaTexto txtObj, 0, fgDetalle
Else
   If Val(Format(fgDetalle.TextMatrix(fgDetalle.Rows - 1, 8), gcFormDato)) <> 0 And _
      Len(Format(fgDetalle.TextMatrix(fgDetalle.Rows - 1, 2), gcFormDato)) > 0 Then
      AdicionaRow fgDetalle, fgDetalle.Rows
      If lbBienes Then
         fgDetalle.TextMatrix(fgDetalle.row, 11) = "B"
      Else
         fgDetalle.TextMatrix(fgDetalle.row, 11) = "S"
      End If
      EnfocaTexto txtObj, 0, fgDetalle
   Else
      If fgDetalle.Enabled Then
         fgDetalle.SetFocus
      End If
   End If
End If
End Sub

Private Sub mnuAtender_Click()
Dim nPos As Integer, nSalto As Integer
nSalto = IIf(fgDetalle.row < fgDetalle.RowSel, 1, -1)
For nPos = fgDetalle.row To fgDetalle.RowSel Step nSalto
    fgDetalle.TextMatrix(nPos, 6) = fgDetalle.TextMatrix(nPos, 5)
Next
SumasDoc
End Sub

Private Sub mnuEliminar_Click()
If fgDetalle.TextMatrix(fgDetalle.row, 0) <> "" Then
   EliminaCuenta fgDetalle.TextMatrix(fgDetalle.row, 2), fgDetalle.TextMatrix(fgDetalle.row, 0)
   SumasDoc
   If fgDetalle.Enabled Then
      fgDetalle.SetFocus
   End If
End If
End Sub

Private Sub EliminaCuenta(sCod As String, nItem As Integer)
EliminaRow fgDetalle, fgDetalle.row
EliminaFgObj nItem
If Len(fgDetalle.TextMatrix(1, 2)) > 0 Then
   RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.row, 0)
End If
End Sub

Private Sub mnuNoAtender_Click()
Dim nPos As Integer, nSalto As Integer
nSalto = IIf(fgDetalle.row < fgDetalle.RowSel, 1, -1)
For nPos = fgDetalle.row To fgDetalle.RowSel Step nSalto
    fgDetalle.TextMatrix(nPos, 6) = ""
Next
SumasDoc
End Sub

Private Sub ActualizaFG(nItem As Integer)
fgDetalle.TextMatrix(nItem, 2) = sObjCod
fgDetalle.TextMatrix(nItem, 3) = sObjDesc
fgDetalle.TextMatrix(nItem, 4) = sObjUnid
fgDetalle.TextMatrix(nItem, 9) = sCtaCod
fgDetalle.TextMatrix(nItem, 10) = sCtaProvis
If fgDetalle.TextMatrix(nItem, 12) = "" Then
   fgDetalle.TextMatrix(nItem, 12) = txtPlazo
End If
'ALPA 20091110***************************************
fgDetalle.TextMatrix(nItem, 1) = sAgeCodDes
'****************************************************
fgDetalle.col = 5
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub cmdDocumentoPDF_Click() 'PASIERS0772014
    Dim oDoc As New cPDF
    Dim sPreNombre As String
    Dim lsCorreos As String
    Dim lsCabecera As String
    Dim ologImp As DLogGeneral
    Dim NtamanioFuente As Integer
    Dim nTamanioFuenteMin As Integer
    Dim nNumLineas As Integer
    Dim nCantCarac As Integer
    Set ologImp = New DLogGeneral
    If chkIGV.value = 1 And chkRenta4.value = 1 Then
        MsgBox "Solo debe elegir 1 concepto IGV / Renta 4ta", vbInformation, "Aviso"
        Exit Sub
    End If
    lsCorreos = LeeConstSistema(451)
    gdFecha = txtFecha 'PASI20150211
    lsCabecera = ""

    If Right(cmbTipoOC.Text, 1) = "D" Then
        sPreNombre = "DIRECTA"
    ElseIf Right(cmbTipoOC.Text, 1) = "P" Then
        sPreNombre = "PROCESO"
    End If
    
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Administrativo"
    oDoc.Producer = gsNomCmac
    oDoc.Subject = "ORDEN"
    oDoc.Title = "ORDEN"
    
    If Not oDoc.PDFCreate(App.path & "\Spooler\Orden_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If

    'Cabecera
    'odoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    'odoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    
    oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    
     NtamanioFuente = 8
     nTamanioFuenteMin = 8
    oDoc.NewPage A4_Vertical
    oDoc.WTextBox 43, 40, 15, 500, "CMAC MAYNAS S.A.", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    Dim lineabase As Integer
    lineabase = 43
   
    If lbBienes Then
        If Right(cmbTipoOC.Text, 1) = "D" Or Right(cmbTipoOC.Text, 1) = "P" Then
            oDoc.WTextBox lineabase, 40, 15, 520, "ORDEN DE COMPRA " & "- " & sPreNombre, "F2", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack
            oDoc.WTextBox lineabase + 10, 40, 15, 500, "RUC: 20103845328", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
            oDoc.WTextBox lineabase + 10, 405, 15, 150, gcDocNro, "F2", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack
        Else
            oDoc.WTextBox lineabase, 40, 15, 520, "ORDEN DE COMPRA ", "F2", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack
            oDoc.WTextBox lineabase + 10, 40, 15, 500, "RUC: 20103845328", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
            oDoc.WTextBox lineabase + 10, 405, 15, 150, gcDocNro, "F2", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack
        End If
    Else
            oDoc.WTextBox lineabase, 40, 15, 520, "ORDEN DE SERVICIO ", "F2", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack
            oDoc.WTextBox lineabase + 10, 40, 15, 500, "RUC: 20103845328", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
            oDoc.WTextBox lineabase + 10, 405, 15, 150, gcDocNro, "F2", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack
    End If
    
    oDoc.WTextBox lineabase + 20, 40, 15, 300, "Mov : " & ologImp.ObtieneMovxImpresionOrden(lsMovNro), "F1", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 20, 413, 15, 140, "Iquitos : " & JIZQ(ArmaFecha(gdFecha), 40), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack 'PASI20150211 Cambio gdFecsis x gdFecha

    'end

    'Datos Proveedor
    gdFecha = txtFecha
    Dim rs1 As ADODB.Recordset
    Dim lnRucProv As String
    Set rs1 = New ADODB.Recordset
    Set rs1 = GetProveedorRUC(txtPersona.Text)
    If Not (rs1.EOF And rs1.BOF) Then
        lnRucProv = rs1!cPersIDnro
    End If
    oDoc.WTextBox lineabase + 30, 40, 15, 600, "PROVEEDOR : " & JIZQ(txtProvNom.Text, 45), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 40, 40, 15, 25, "RUC : ", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 40, 65, 15, 100, JIZQ(lnRucProv, 12), "F1", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 50, 40, 15, 70, "DIRECCION : ", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 50, 110, 15, 500, JIZQ(txtProvDir.Text, 45), "F1", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    'odoc.WTextBox lineabase + 60, 40, 15, 20, "Telef.: ", "F2", 7, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 60, 40, 15, 100, "Telef.: " & JIZQ(txtProvTele.Text, 12), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    'end
    'Detalle Cabecera
    oDoc.WTextBox lineabase + 70, 40, 15, 550, "Sírvase atender de acuerdo al siguiente detalle:", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 80, 40, 15, 570, String(118, "_"), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 90, 40, 15, 70, "Agencia", "F2", NtamanioFuente, AlignH.hCenter, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 90, 110, 15, 5, "|", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 90, 115, 15, 45, "Cantidad", "F2", NtamanioFuente, AlignH.hCenter, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 90, 160, 15, 5, "|", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 90, 165, 15, 40, "Unidad", "F2", NtamanioFuente, AlignH.hCenter, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 90, 205, 15, 5, "|", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 90, 210, 15, 230, "Descripción", "F2", NtamanioFuente, AlignH.hCenter, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 90, 425, 15, 5, "|", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 90, 435, 15, 60, "P.Vta.Unit", "F2", NtamanioFuente, AlignH.hCenter, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 90, 495, 15, 5, "|", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 90, 500, 15, 60, "P.Vta.Total", "F2", NtamanioFuente, AlignH.hCenter, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 90, 560, 15, 5, "|", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    oDoc.WTextBox lineabase + 92, 40, 15, 570, String(118, "_"), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    'end Detalle Cabecera
    
    'Detalle Bienes/Servicios
    Dim lsCantidadCad As String
    Dim lsTotalCad As String

    Dim I, J, nlineas, varpos, lineaBasePdf As Integer
    Dim lnBaseDetalle As Integer
    Dim varDesc As String
    Dim varTope As Integer
    
    I = 1
    varTope = 800
    lnBaseDetalle = 145
    varpos = 1
    lineaBasePdf = 145
    nNumLineas = 45
     nCantCarac = 70
    Do While Not fgDetalle.Rows - 1
        If fgDetalle.TextMatrix(I, 2) = "" Then
            Exit Do
        End If
        lsCantidadCad = fgDetalle.TextMatrix(I, 6)
        lsTotalCad = fgDetalle.TextMatrix(I, 8)
        If lbBienes Then 'Detalle Bienes/Servicios
            If I = 1 Then
                oDoc.WTextBox lnBaseDetalle, 40, 15, 80, Replace(Replace(GetAgencias(fgDetalle.TextMatrix(I, 1)), "Agencia", "Ag."), "Oficina", "Of."), "F1", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack 'Agencia
                oDoc.WTextBox lnBaseDetalle, 115, 15, 40, fgDetalle.TextMatrix(I, 5), "F1", NtamanioFuente, AlignH.hCenter, vTop, vbBlack, 0, vbBlack  'Cantidad
                oDoc.WTextBox lnBaseDetalle, 165, 15, 40, fgDetalle.TextMatrix(I, 4), "F1", NtamanioFuente, AlignH.hCenter, vTop, vbBlack, 0, vbBlack  'Unidad
                oDoc.WTextBox lnBaseDetalle, 430, 15, 60, fgDetalle.TextMatrix(I, 6), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'Pv Unit.
                oDoc.WTextBox lnBaseDetalle, 495, 15, 55, fgDetalle.TextMatrix(I, 8), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'Pv Total
            'End If
            'nLineas = Round((Len(fgDetalle.TextMatrix(i, 3)) / nNumLineas) + 0.4)
            nlineas = Round((Len(fgDetalle.TextMatrix(I, 3)) / nNumLineas) + 0.4)
            varDesc = fgDetalle.TextMatrix(I, 3)
            'varDesc = JustificaTextoCadenaPASI((fgDetalle.TextMatrix(i, 3)), nCantCarac, 1) 'otiginal
            'nLineas = Round((Len(varDesc) / nNumLineas) + 0.4) 'original
            oDoc.WTextBox lnBaseDetalle, 210, 15, 200, varDesc, "F1", NtamanioFuente, AlignH.hjustify, vTop, vbBlack, 0, vbBlack    'Descripcion
            lnBaseDetalle = lnBaseDetalle + nlineas * 7
            Else
                'nLineas = Round((Len(fgDetalle.TextMatrix(i, 3)) / nNumLineas) + 0.4)
                nlineas = Round((Len(fgDetalle.TextMatrix(I, 3)) / nNumLineas) + 0.4)
                varDesc = fgDetalle.TextMatrix(I, 3)
                'varDesc = JustificaTextoCadenaPASI((fgDetalle.TextMatrix(i, 3)), nCantCarac, 1) 'otiginal
                'nLineas = Round((Len(varDesc) / nNumLineas) + 0.4) 'original
                lineaBasePdf = lnBaseDetalle + nlineas * 7
                If lineaBasePdf >= varTope Then
                    oDoc.NewPage A4_Vertical
                   lnBaseDetalle = 43
                End If
                oDoc.WTextBox lnBaseDetalle, 40, 15, 80, Replace(Replace(GetAgencias(fgDetalle.TextMatrix(I, 1)), "Agencia", "Ag."), "Oficina", "Of."), "F1", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack 'Agencia
                oDoc.WTextBox lnBaseDetalle, 115, 15, 40, fgDetalle.TextMatrix(I, 5), "F1", NtamanioFuente, AlignH.hCenter, vTop, vbBlack, 0, vbBlack  'Cantidad
                oDoc.WTextBox lnBaseDetalle, 165, 15, 40, fgDetalle.TextMatrix(I, 4), "F1", NtamanioFuente, AlignH.hCenter, vTop, vbBlack, 0, vbBlack  'Unidad
                oDoc.WTextBox lnBaseDetalle, 430, 15, 60, fgDetalle.TextMatrix(I, 6), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'Pv Unit.
                oDoc.WTextBox lnBaseDetalle, 495, 15, 55, fgDetalle.TextMatrix(I, 8), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'Pv Total
                oDoc.WTextBox lnBaseDetalle, 210, 15, 200, varDesc, "F1", NtamanioFuente, AlignH.hjustify, vTop, vbBlack, 0, vbBlack    'Descripcion
                lnBaseDetalle = lnBaseDetalle + nlineas * 7
            End If
        Else
            If I = 1 Then
                oDoc.WTextBox lnBaseDetalle, 40, 15, 80, Replace(Replace(GetAgencias(fgDetalle.TextMatrix(I, 1)), "Agencia", "Ag."), "Oficina", "Of."), "F1", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack 'Agencia
                oDoc.WTextBox lnBaseDetalle, 115, 15, 40, fgDetalle.TextMatrix(I, 5), "F1", NtamanioFuente, AlignH.hCenter, vTop, vbBlack, 0, vbBlack  'Cantidad
                oDoc.WTextBox lnBaseDetalle, 165, 15, 40, fgDetalle.TextMatrix(I, 4), "F1", NtamanioFuente, AlignH.hCenter, vTop, vbBlack, 0, vbBlack  'Unidad
                oDoc.WTextBox lnBaseDetalle, 430, 15, 60, fgDetalle.TextMatrix(I, 6), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'Pv Unit.
                oDoc.WTextBox lnBaseDetalle, 495, 15, 55, fgDetalle.TextMatrix(I, 8), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'Pv Total
            'End If
                'nLineas = Round((Len(fgDetalle.TextMatrix(i, 3)) / nNumLineas) + 0.4)
                nlineas = Round((Len(fgDetalle.TextMatrix(I, 3)) / nNumLineas) + 0.4)
                varDesc = fgDetalle.TextMatrix(I, 3)
                'varDesc = JustificaTextoCadenaPASI((fgDetalle.TextMatrix(i, 3)), nCantCarac, 1) 'otiginal
                'nLineas = Round((Len(varDesc) / nNumLineas) + 0.4) 'original
                'nLineas = Round((Len(varDesc) / nNumLineas) + 0.4)
                oDoc.WTextBox lnBaseDetalle, 210, 15, 200, varDesc, "F1", NtamanioFuente, AlignH.hjustify, vTop, vbBlack, 0, vbBlack    'Descripcion
                lnBaseDetalle = lnBaseDetalle + nlineas * 7
            Else
                'nLineas = Round((Len(fgDetalle.TextMatrix(i, 3)) / nNumLineas) + 0.4)
                nlineas = Round((Len(fgDetalle.TextMatrix(I, 3)) / nNumLineas) + 0.4)
                varDesc = fgDetalle.TextMatrix(I, 3)
                'varDesc = JustificaTextoCadenaPASI((fgDetalle.TextMatrix(i, 3)), nCantCarac, 1) 'otiginal
                'nLineas = Round((Len(varDesc) / nNumLineas) + 0.4) 'original
                lineaBasePdf = lnBaseDetalle + nlineas * 7
                If lineaBasePdf >= varTope Then
                    oDoc.NewPage A4_Vertical
                   lnBaseDetalle = 43
                End If
                oDoc.WTextBox lnBaseDetalle, 40, 15, 80, Replace(Replace(GetAgencias(fgDetalle.TextMatrix(I, 1)), "Agencia", "Ag."), "Oficina", "Of."), "F1", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack 'Agencia
                oDoc.WTextBox lnBaseDetalle, 115, 15, 40, fgDetalle.TextMatrix(I, 5), "F1", NtamanioFuente, AlignH.hCenter, vTop, vbBlack, 0, vbBlack  'Cantidad
                oDoc.WTextBox lnBaseDetalle, 165, 15, 40, fgDetalle.TextMatrix(I, 4), "F1", NtamanioFuente, AlignH.hCenter, vTop, vbBlack, 0, vbBlack  'Unidad
                oDoc.WTextBox lnBaseDetalle, 430, 15, 60, fgDetalle.TextMatrix(I, 6), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'Pv Unit.
                oDoc.WTextBox lnBaseDetalle, 495, 15, 55, fgDetalle.TextMatrix(I, 8), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'Pv Total
                oDoc.WTextBox lnBaseDetalle, 210, 15, 200, varDesc, "F1", NtamanioFuente, AlignH.hjustify, vTop, vbBlack, 0, vbBlack    'Descripcion
                lnBaseDetalle = lnBaseDetalle + nlineas * 7
            End If
        
            
        End If
        
        
        I = I + 1
        lnBaseDetalle = lnBaseDetalle + 7
    Loop
    'end Detalle Bienes/Servicios
    'Totales
        'lnBaseDetalle = lnBaseDetalle - 10
        oDoc.WTextBox lnBaseDetalle + 5, 40, 15, 570, String(118, "_"), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        Dim lsMonto As String * 74
        lsMonto = Trim(CStr(ConvNumLet(nVal(txtTot.Text))))
        Dim nSubRenta As Currency
        Dim nRenta4 As Currency
        Dim nSubt As Currency
        Dim nIGVT As Currency
        
        If lnBaseDetalle >= varTope Then
            oDoc.NewPage A4_Vertical
            lnBaseDetalle = 43
        End If
        
        lnBaseDetalle = lnBaseDetalle + 20
        If lbBienes Then
            nSubt = Format(((txtTot.Text) / (1 + gnIGV)), "0.00")
            nIGVT = txtTot.Text - nSubt
            If chkIGV.value = 1 Then
                    If lnBaseDetalle >= varTope Then
                        oDoc.NewPage A4_Vertical
                        lnBaseDetalle = 43
                    End If
                    '''oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack  'Vv Total Text 'MARG ERS044-2016
                    oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack  'Vv Total Text
                    oDoc.WTextBox lnBaseDetalle, 495, 15, 55, JDER(Format(nSubt, gcFormView), 18), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'Vv Total Monto
                    lnBaseDetalle = lnBaseDetalle + 20
                    If lnBaseDetalle >= varTope Then
                        oDoc.NewPage A4_Vertical
                        lnBaseDetalle = 43
                    End If
                    oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "I.G.V. ", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack  'IGV Text
                    oDoc.WTextBox lnBaseDetalle, 495, 15, 55, JDER(Format(nIGVT, gcFormView), 27), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'IGV Monto
            Else
                    If lnBaseDetalle >= varTope Then
                        oDoc.NewPage A4_Vertical
                        lnBaseDetalle = 43
                    End If
                    '''oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack 'Vv Total Text 'MARG ERS044-2016
                    oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack 'Vv Total Text
                    oDoc.WTextBox lnBaseDetalle, 495, 15, 55, JDER(Format(txtTot.Text, gcFormView), 18), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'Vv Total Monto
                     lnBaseDetalle = lnBaseDetalle + 20
                    If lnBaseDetalle >= varTope Then
                        oDoc.NewPage A4_Vertical
                        lnBaseDetalle = 43
                    End If
            End If
            oDoc.WTextBox lnBaseDetalle, 500, 15, 60, "__________", "F2", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack
            lnBaseDetalle = lnBaseDetalle + 20
            If lnBaseDetalle >= varTope Then
                oDoc.NewPage A4_Vertical
                lnBaseDetalle = 43
            End If
            oDoc.WTextBox lnBaseDetalle, 40, 15, 300, "SON: " & lsMonto, "F1", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
            '''oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack 'Pv Total Monto 'MARG ERS04--2016
            oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack 'Pv Total Monto
            oDoc.WTextBox lnBaseDetalle, 495, 15, 55, JDER(txtTot.Text, 18), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack
        Else
            If chkRenta4.value = 1 Then
                'nSubRenta = Format(((CCur(txtTot.Text)) / 8), "0.00") 'Comentado PASI20150211
                nSubRenta = Format(((CCur(txtTot.Text)) * 0.08), "0.00") 'PASI20151102 Cambio la Retencion de 10 a 8 segun TIC1502100009
                nRenta4 = txtTot.Text - nSubRenta
                If lnBaseDetalle >= varTope Then
                        oDoc.NewPage A4_Vertical
                        lnBaseDetalle = 43
                End If
                '''oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack  'Vv Total Text 'MARG ERS044-2016
                oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack  'Vv Total Text 'MARG ERS044-2016
                oDoc.WTextBox lnBaseDetalle, 495, 15, 55, JDER(Format(txtTot.Text, gcFormView), 20), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'Vv Total Monto
                lnBaseDetalle = lnBaseDetalle + 10
                If lnBaseDetalle >= varTope Then
                    oDoc.NewPage A4_Vertical
                    lnBaseDetalle = 43
                End If
                oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "Ret. 4ta. ", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack   'ret 4 Text
                oDoc.WTextBox lnBaseDetalle, 495, 15, 55, JDER(Format(nSubRenta, gcFormView), 26), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'ret 4 Monto
                oDoc.WTextBox lnBaseDetalle, 500, 15, 60, "__________", "F2", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack
                lnBaseDetalle = lnBaseDetalle + 10
                If lnBaseDetalle >= varTope Then
                    oDoc.NewPage A4_Vertical
                    lnBaseDetalle = 43
                End If
                '''oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack 'Pv Total Monto 'MARG ERS044-2016
                oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack 'Pv Total Monto 'MARG ERS044-2016
                oDoc.WTextBox lnBaseDetalle, 495, 15, 55, JDER(Format(nRenta4, gcFormView), 20), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack
                oDoc.WTextBox lnBaseDetalle, 40, 15, 300, "SON: " & lsMonto, "F1", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
            Else
                If chkIGV.value = 1 Then
                    nSubt = Format(((txtTot.Text) / (1 + gnIGV)), "0.00")
                    nIGVT = txtTot.Text - nSubt
                    If lnBaseDetalle >= varTope Then
                        oDoc.NewPage A4_Vertical
                        lnBaseDetalle = 43
                    End If
                    '''oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack  'Vv Total Text 'marg ers044-2016
                    oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack  'Vv Total Text 'marg ers044-2016
                    oDoc.WTextBox lnBaseDetalle, 495, 15, 55, JDER(Format(nSubt, gcFormView), 18), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'Vv Total Monto
                    lnBaseDetalle = lnBaseDetalle + 10
                    If lnBaseDetalle >= varTope Then
                        oDoc.NewPage A4_Vertical
                        lnBaseDetalle = 43
                    End If
                    oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "I.G.V. ", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack  'IGV Text
                    oDoc.WTextBox lnBaseDetalle, 495, 15, 55, JDER(Format(nIGVT, gcFormView), 27), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'IGV Monto
                    oDoc.WTextBox lnBaseDetalle, 500, 15, 60, "__________", "F2", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack
                    lnBaseDetalle = lnBaseDetalle + 10
                    If lnBaseDetalle >= varTope Then
                        oDoc.NewPage A4_Vertical
                        lnBaseDetalle = 43
                    End If
                    '''oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack 'Pv Total Monto 'marg ers044-2016
                    oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack 'Pv Total Monto 'marg ers044-2016
                    
                    oDoc.WTextBox lnBaseDetalle, 495, 15, 55, JDER(txtTot.Text, 20), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack
                    oDoc.WTextBox lnBaseDetalle, 40, 15, 300, "SON: " & lsMonto, "F1", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
                Else
                    If lnBaseDetalle >= varTope Then
                        oDoc.NewPage A4_Vertical
                        lnBaseDetalle = 43
                    End If
                    '''oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack  'Vv Total Text 'MARG ERS044-2016
                    oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack  'Vv Total Text 'MARG ERS044-2016
                    oDoc.WTextBox lnBaseDetalle, 495, 15, 55, JDER(txtTot.Text, 18), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack  'Vv Total Monto
                    oDoc.WTextBox lnBaseDetalle, 500, 15, 60, "__________", "F2", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack
                    lnBaseDetalle = lnBaseDetalle + 10
                    If lnBaseDetalle >= varTope Then
                        oDoc.NewPage A4_Vertical
                        lnBaseDetalle = 43
                    End If
                    '''oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack 'Pv Total Monto 'MARG ERS044-2015
                    oDoc.WTextBox lnBaseDetalle, 420, 15, 70, "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. "), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack 'Pv Total Monto 'MARG ERS044-2015
                    oDoc.WTextBox lnBaseDetalle, 495, 15, 55, JDER(txtTot.Text, 20), "F1", NtamanioFuente, AlignH.hRight, vTop, vbBlack, 0, vbBlack
                    oDoc.WTextBox lnBaseDetalle, 40, 15, 300, "SON: " & lsMonto, "F1", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
                End If
            End If
        End If
    'end Totales
    'Observaciones
        oDoc.WTextBox lnBaseDetalle + 10, 40, 15, 570, String(118, "_"), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        lnBaseDetalle = lnBaseDetalle + 20
        If lnBaseDetalle >= varTope Then
            oDoc.NewPage A4_Vertical
            lnBaseDetalle = 43
        End If
        oDoc.WTextBox lnBaseDetalle, 40, 15, 500, "Concepto y Observaciones:", "F2", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        lnBaseDetalle = lnBaseDetalle + 5
        If lnBaseDetalle >= varTope Then
            oDoc.NewPage A4_Vertical
            lnBaseDetalle = 43
        End If
        
        Dim sGlosa As String
        Dim nLineasObs As String
        sGlosa = Trim(txtMovDesc.Text)
        
        sGlosa = JustificaTextoCadenaPASI((sGlosa), 360, 1)
        nLineasObs = Round((Len(sGlosa) / 370) + 0.4)
        
        lnBaseDetalle = lnBaseDetalle + nLineasObs * 5
        If lnBaseDetalle >= varTope Then
            oDoc.NewPage A4_Vertical
           lnBaseDetalle = 43
        End If
        oDoc.WTextBox lnBaseDetalle, 40, 15, 530, sGlosa, "F1", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        oDoc.WTextBox lnBaseDetalle + 10, 40, 15, 570, String(118, "_"), "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        lnBaseDetalle = lnBaseDetalle + 20
        If lnBaseDetalle >= varTope Then
            oDoc.NewPage A4_Vertical
           lnBaseDetalle = 43
        End If
        'odoc.WTextBox lnBaseDetalle, 40, 15, 600, JustificaTextoCadenaOrdenCompra(("Los precios incluyen todos los tributos, seguros, transportes, inspecciones, pruebas y cualquier otro concepto que pueda incidir sobre el costo del(os) bien(es)."), 200, 1), "F1", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        oDoc.WTextBox lnBaseDetalle, 40, 15, 520, "Los precios incluyen todos los tributos, seguros, transportes, inspecciones, pruebas y cualquier otro concepto que pueda incidir sobre el costo del(os) bien(es). Adjuntar copia de la orden de  " & IIf(lbBienes, "compra ", "servicio ") & "a la factura.", "F1", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        lnBaseDetalle = lnBaseDetalle + 20
        If lnBaseDetalle >= varTope Then
            oDoc.NewPage A4_Vertical
           lnBaseDetalle = 43
        End If
'        odoc.WTextBox lnBaseDetalle, 40, 15, 600, JustificaTextoCadenaOrdenCompra(("Adjuntar copia de la orden de  " & IIf(lbBienes, "compra ", "servicio ") & "a la factura."), 200, 1), "F1", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
         lnBaseDetalle = lnBaseDetalle + 10
        If lnBaseDetalle >= varTope Then
            oDoc.NewPage A4_Vertical
           lnBaseDetalle = 43
        End If
        oDoc.WTextBox lnBaseDetalle, 110, 15, 600, " FACTURAR A NOMBRE DE: ", "F2", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        oDoc.WTextBox lnBaseDetalle, 230, 15, 600, "CMAC MAYNAS S.A.; RUC: 20103845328; Jr. Próspero N° 791 - Iquitos.", "F1", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        lnBaseDetalle = lnBaseDetalle + 10
        If lnBaseDetalle >= varTope Then
            oDoc.NewPage A4_Vertical
           lnBaseDetalle = 43
        End If
        oDoc.WTextBox lnBaseDetalle, 40, 15, 110, "PLAZO DE ENTREGA: ", "F2", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        oDoc.WTextBox lnBaseDetalle, 150, 15, 50, txtPlazo, "F1", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        oDoc.WTextBox lnBaseDetalle, 230, 15, 80, "FORMA DE PAGO:   ", "F2", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        oDoc.WTextBox lnBaseDetalle, 310, 15, 50, Trim(Left(cboFormaPago, Len(cboFormaPago) - 1)), "F1", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        oDoc.WTextBox lnBaseDetalle, 420, 15, 80, "GARANTIA:   ", "F2", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        oDoc.WTextBox lnBaseDetalle, 480, 15, 40, txtDiasAtraso.Text & " meses", "F1", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        lnBaseDetalle = lnBaseDetalle + 10
        If lnBaseDetalle >= varTope Then
            oDoc.NewPage A4_Vertical
           lnBaseDetalle = 43
        End If
'        odoc.WTextBox lnBaseDetalle, 40, 15, 500, "IMPORTANTE:", "F2", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
'        lnBaseDetalle = lnBaseDetalle + 10
'        If lnBaseDetalle >= varTope Then
'            odoc.NewPage A4_Vertical
'           lnBaseDetalle = 43
'        End If
        If lbBienes Then
            Dim lsObs As String
            Dim nlinea As Integer
             oDoc.WTextBox lnBaseDetalle, 40, 15, 500, "IMPORTANTE:", "F2", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
            lnBaseDetalle = lnBaseDetalle + 10
            If lnBaseDetalle >= varTope Then
                oDoc.NewPage A4_Vertical
               lnBaseDetalle = 43
            End If
            lsObs = "1. El número de esta Orden debera consignarse en facturas, guías remisión, embalaje y correspondencia respectiva."
            lsObs = lsObs & " Asimismo el Contratista debe acusar recibo de esta Orden inmediatemente después de su recepción."
            oDoc.WTextBox lnBaseDetalle, 40, 15, 485, lsObs, "F1", nTamanioFuenteMin, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
            lnBaseDetalle = lnBaseDetalle + 25
            If lnBaseDetalle >= varTope Then
                oDoc.NewPage A4_Vertical
                lnBaseDetalle = 43
            End If
            lsObs = "2. La CMAC MAYNAS S.A. se reserva el derecho de devolver los bienes, si no cumplen con las especificaciones "
            lsObs = lsObs & "mínimas requeridas, y de acuerdo a lo ofertado por el Contratista, así como de anular la Orden de Compra,"
            lsObs = lsObs & " según el procedimiento establecido en el reglamento de adquisiciones y contrataciones para el sistema CMAC."
            oDoc.WTextBox lnBaseDetalle, 40, 15, 485, lsObs, "F1", nTamanioFuenteMin, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
            lnBaseDetalle = lnBaseDetalle + 35
            If lnBaseDetalle >= varTope Then
                oDoc.NewPage A4_Vertical
                lnBaseDetalle = 43
            End If
            lsObs = "3. El Contratista se compromete a cumplir las obligaciones que le corresponden, bajo sanción de quedar inhabilitado para contratar con el Estado en caso de incumplimiento."
            oDoc.WTextBox lnBaseDetalle, 40, 15, 485, lsObs, "F1", nTamanioFuenteMin, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
            lnBaseDetalle = lnBaseDetalle + 25
            If lnBaseDetalle >= varTope Then
                oDoc.NewPage A4_Vertical
                lnBaseDetalle = 43
            End If
            '[LARI 20200331 *************************************
            lsObs = "4. En caso de retraso injustificado en la entrega de los bienes,  la CAJA aplicará al contratista una penalidad por cada día de atraso, hasta por un monto máximo equivalente al Diez por ciento (10%) del monto de la ODC vigente."
            oDoc.WTextBox lnBaseDetalle, 40, 15, 485, lsObs, "F1", nTamanioFuenteMin, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
            lnBaseDetalle = lnBaseDetalle + 20
            
            lsObs = "Cálculo de penalidad diaria=(0.10 x Monto)/(F x Plazo en días)"
            oDoc.WTextBox lnBaseDetalle, 40, 15, 485, lsObs, "F1", nTamanioFuenteMin, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
            lnBaseDetalle = lnBaseDetalle + 10
            
            lsObs = "-Para plazos menores o iguales a sesenta (60) días, para bienes y servicios F=0.40"
            oDoc.WTextBox lnBaseDetalle, 40, 15, 485, lsObs, "F1", nTamanioFuenteMin, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
            lnBaseDetalle = lnBaseDetalle + 10
            
            lsObs = "-Para plazos mayores a sesenta (60) días, para bienes y servicio F=0.25"
            oDoc.WTextBox lnBaseDetalle, 40, 15, 485, lsObs, "F1", nTamanioFuenteMin, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
            lnBaseDetalle = lnBaseDetalle + 20
            '****************************************************]
            
            
            If lnBaseDetalle >= varTope Then
                oDoc.NewPage A4_Vertical
                lnBaseDetalle = 43
            End If
            lsObs = "5. Para Agilizar el pago a los proveedores se deberá alcanzar número de cuenta de ahorros en la Caja Maynas o en su defecto indicar número de cuenta en el BCP."
            oDoc.WTextBox lnBaseDetalle, 40, 15, 485, lsObs, "F1", nTamanioFuenteMin, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
            lnBaseDetalle = lnBaseDetalle + 25
            If lnBaseDetalle >= varTope Then
                oDoc.NewPage A4_Vertical
                lnBaseDetalle = 43
            End If
            lsObs = "6. El día de pago a proveedores se realiza todos los viernes, previa conformidad del área usuaria, caso contrario de tener observaciones este se realizaran en 15 días hábiles de"
            lsObs = lsObs & " presentado el comprobante."
            oDoc.WTextBox lnBaseDetalle, 40, 15, 485, lsObs, "F1", nTamanioFuenteMin, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
        Else
            '[LARI 20200331 ***********************************
             oDoc.WTextBox lnBaseDetalle, 40, 15, 500, "IMPORTANTE:", "F2", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
            lnBaseDetalle = lnBaseDetalle + 10
            If lnBaseDetalle >= varTope Then
                oDoc.NewPage A4_Vertical
               lnBaseDetalle = 43
            End If
            
            lsObs = "1. En caso de retraso injustificado en la entrega de los bienes,  la CAJA aplicará al contratista una penalidad por cada día de atraso, hasta por un monto máximo equivalente al Diez por ciento (10%) del monto de la ODC vigente."
            oDoc.WTextBox lnBaseDetalle, 40, 15, 485, lsObs, "F1", nTamanioFuenteMin, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
            lnBaseDetalle = lnBaseDetalle + 20
            
            lsObs = "   Cálculo de penalidad diaria=(0.10 x Monto)/(F x Plazo en días)"
            oDoc.WTextBox lnBaseDetalle, 40, 15, 485, lsObs, "F1", nTamanioFuenteMin, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
            lnBaseDetalle = lnBaseDetalle + 10
            
            lsObs = "   -Para plazos menores o iguales a sesenta (60) días, para bienes y servicios F=0.40"
            oDoc.WTextBox lnBaseDetalle, 40, 15, 485, lsObs, "F1", nTamanioFuenteMin, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
            lnBaseDetalle = lnBaseDetalle + 10
            
            lsObs = "   -Para plazos mayores a sesenta (60) días, para bienes y servicio F=0.25"
            oDoc.WTextBox lnBaseDetalle, 40, 15, 485, lsObs, "F1", nTamanioFuenteMin, AlignH.hjustify, vTop, vbBlack, 0, vbBlack
            'lnBaseDetalle = lnBaseDetalle + 10
            If lnBaseDetalle >= varTope Then
                oDoc.NewPage A4_Vertical
                lnBaseDetalle = 43
            End If
            '************************************************]
        End If
        If lnBaseDetalle + 150 >= varTope Then
            oDoc.NewPage A4_Vertical
            lnBaseDetalle = 43
        End If
        lnBaseDetalle = lnBaseDetalle + 100
        'odoc.WTextBox lnBaseDetalle + 5, 40, 15, 550, String(95, "_"), "F2", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
'        lnBaseDetalle = lnBaseDetalle + 15
'        odoc.WTextBox lnBaseDetalle, 40, 15, 100, "AREA USUARIA", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
'        odoc.WTextBox lnBaseDetalle, 200, 15, 100, "LOGISTICA", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
'        odoc.WTextBox lnBaseDetalle, 350, 15, 100, "GERENCIA", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
'        odoc.WTextBox lnBaseDetalle, 500, 15, 100, "PROVEEDOR", "F2", NtamanioFuente, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        
        oDoc.WTextBox lnBaseDetalle, 40, 15, 570, String(118, "_"), "F2", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        lnBaseDetalle = lnBaseDetalle + 10
        oDoc.WTextBox lnBaseDetalle, 40, 15, 100, "AREA USUARIA", "F2", 8, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        oDoc.WTextBox lnBaseDetalle, 275, 15, 100, "LOGISTICA", "F2", 8, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        'odoc.WTextBox lnBaseDetalle, 350, 15, 100, "GERENCIA", "F2", 8, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        oDoc.WTextBox lnBaseDetalle, 500, 15, 100, "PROVEEDOR", "F2", 8, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        
        If lnBaseDetalle + 150 >= varTope Then
            oDoc.NewPage A4_Vertical
            lnBaseDetalle = 43
        End If
        
        oDoc.WTextBox varTope - 20, 40, 15, 570, String(118, "_"), "F2", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        oDoc.WTextBox varTope - 10, 40, 15, 550, "Cualquier Consulta: Teléfono (065) 58-1770 Anexos 1412 - 1413 / Fax Anexo 1414", "F2", 8, AlignH.hCenter, vTop, vbBlack, 0, vbBlack
        lsCorreos = LeeConstSistema(451)
        oDoc.WTextBox varTope, 40, 15, 560, "Correos: " + lsCorreos, "F2", 7, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
        oDoc.WTextBox varTope + 5, 40, 15, 570, String(118, "_"), "F2", nTamanioFuenteMin, AlignH.hLeft, vTop, vbBlack, 0, vbBlack
    'end Observaciones
        
        'ARLO 20170125
        If (cmbmoneda.ListIndex = 1) And sDocDesc = "Orden de Compra" Then
        gsOpeCod = LogPistasImpresionOrdenComprasSoles
        ElseIf (cmbmoneda.ListIndex <> 1) And sDocDesc = "Orden de Compra" Then
        gsOpeCod = LogPistasImpresionOrdenComprasDolares
        ElseIf (cmbmoneda.ListIndex = 1) And sDocDesc = "Orden de Servicio" Then
        gsOpeCod = LogPistasImpresionOrdenServicioSoles
        Else: gsOpeCod = LogPistasImpresionOrdenServicioDolares
        End If
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Imprimio la " & sDocDesc & " " & sPreNombre & " en " & Mid(cmbmoneda.Text, 1, 7) & " N°: " & gcDocNro
        Set objPista = Nothing
        '***********
        
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Sub cmdDocumento_Click()
Dim oPrevio As clsPrevio
Dim lsCadena As String
Dim sPreNombre As String
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim I As Integer
Dim k As Integer
Dim div As Integer
Dim J As Integer
Dim Desc As String
Dim DescF As String
Dim sGlosa As String
Dim sGlosaF As String
Dim lsCorreos As String 'EJVG20131126

Dim nSubt As Currency
Dim nIGVT As Currency
Dim lineas As Long
Dim lsCabecera As String

If chkIGV.value = 1 And chkRenta4.value = 1 Then
   MsgBox "Solo debe elegir 1 concepto IGV / Renta 4ta", vbInformation, "Aviso"
   Exit Sub
End If

lsCorreos = LeeConstSistema(451) 'EJVG20131126

Set oPrevio = New clsPrevio

lsNegritaOn = oImpresora.gPrnBoldON
lsNegritaOff = oImpresora.gPrnBoldOFF
lsCabecera = ""

If Right(cmbTipoOC.Text, 1) = "D" Then
    sPreNombre = " DIRECTA "
ElseIf Right(cmbTipoOC.Text, 1) = "P" Then
   sPreNombre = " PROCESO "
End If

lsCabecera = lsCabecera & "." & Space(8) & lsNegritaOn & "CMAC MAYNAS S.A." & lsNegritaOff

If lbBienes Then
   If Right(cmbTipoOC.Text, 1) = "D" Or Right(cmbTipoOC.Text, 1) = "P" Then
        lsCabecera = lsCabecera & Space(66) & lsNegritaOn & "ORDEN DE COMPRA " & lsNegritaOff & "- " & sPreNombre & oImpresora.gPrnSaltoLinea
        lsCabecera = lsCabecera & Space(10) & lsNegritaOn & "RUC: 20103845328" & lsNegritaOff
        'lsCabecera = lsCabecera & Space(10) & lsNegritaOn & Space(75) & "ORDEN DE COMPRA " & lsNegritaOff & "- " & sPrenombre & oImpresora.gPrnSaltoLinea
        lsCabecera = lsCabecera & Space(66) & lsNegritaOn & gcDocNro & lsNegritaOff & oImpresora.gPrnSaltoLinea
   Else
        lsCabecera = lsCabecera & Space(66) & lsNegritaOn & "ORDEN DE COMPRA " & lsNegritaOff & oImpresora.gPrnSaltoLinea
        lsCabecera = lsCabecera & Space(10) & lsNegritaOn & "RUC: 20103845328" & lsNegritaOff
        lsCabecera = lsCabecera & Space(66) & lsNegritaOn & gcDocNro & lsNegritaOff & oImpresora.gPrnSaltoLinea
   End If
Else
   lsCabecera = lsCabecera & Space(66) & lsNegritaOn & "ORDEN DE SERVICIO " & lsNegritaOff & oImpresora.gPrnSaltoLinea
   lsCabecera = lsCabecera & Space(10) & lsNegritaOn & "RUC: 20103845328" & lsNegritaOff
   lsCabecera = lsCabecera & Space(66) & gcDocNro & lsNegritaOff & oImpresora.gPrnSaltoLinea
End If
gdFecha = txtFecha
  lsCadena = lsCabecera
  
''lsCadena = lsCadena & Chr$(27) & Chr$(50)   'espaciamiento lineas 1/6 pulg.
''lsCadena = lsCadena & Chr$(27) & Chr$(67) & Chr$(22)  'Longitud de página a 22 líneas'
''lsCadena = lsCadena & Chr$(27) & Chr$(77)   'Tamaño 10 cpi
''lsCadena = lsCadena & Chr$(27) + Chr$(107) + Chr$(0) 'Tipo de Letra Sans Serif

Dim rs1 As ADODB.Recordset
Dim lnRucProv As String
Set rs1 = New ADODB.Recordset

Set rs1 = GetProveedorRUC(txtPersona.Text)
If Not (rs1.EOF And rs1.BOF) Then
        lnRucProv = rs1!cPersIDnro
End If

  lsCadena = lsCadena & Space(10) & lsNegritaOn & "PROVEEDOR : " & JIZQ(txtProvNom.Text, 45) & lsNegritaOff & Space(3) & lsNegritaOn & "RUC:  " & lsNegritaOff & JIZQ(lnRucProv, 12)
  lsCadena = lsCadena & Space(4) & "Iquitos, " & JIZQ(ArmaFecha(gdFecha), 40) & oImpresora.gPrnSaltoLinea 'PASI20150211 Cambio gdFecsis x gdFecha
  lsCadena = lsCadena & Space(10) & lsNegritaOn & "DIRECCION : " & lsNegritaOff & JIZQ(txtProvDir.Text, 45) & Space(3) & lsNegritaOn & "Telef.: " & lsNegritaOff & JIZQ(txtProvTele.Text, 12) & oImpresora.gPrnSaltoLinea

  lsCadena = lsCadena & Space(10) & lsNegritaOn & "Sírvase atender de acuerdo al siguiente detalle" & oImpresora.gPrnSaltoLinea

  lsCadena = lsCadena & Space(10) & String(115, "_") & oImpresora.gPrnSaltoLinea
  lsCadena = lsCadena & Space(10) & Space(1) & "Agencia" & Space(1) & "|" & Space(11) & "Cantidad" & Space(1) & "|" & Space(2) & "Unidad" & Space(2) & "|" & Space(1) & "Descripción" & Space(30) & "|" & Space(1) & "P.Vta.Unit" & " |" & Space(3) & "P.Vta.Total" & Space(1) & "|" & oImpresora.gPrnSaltoLinea
  lsCadena = lsCadena & Space(10) & String(115, "_") & lsNegritaOff & oImpresora.gPrnSaltoLinea

  lineas = 8
  
  I = 1

  Do While Not fgDetalle.Rows - 1
    If fgDetalle.TextMatrix(I, 2) = "" Then
       Exit Do
    End If
    Dim lsCantidadCad As String
    Dim lsTotalCad As String
                
    lsCantidadCad = fgDetalle.TextMatrix(I, 6)
    lsTotalCad = fgDetalle.TextMatrix(I, 8)
  'ALPA 20100409**********************************************
     'Desc = JustificaTextoCadenaOrdenCompra(fgDetalle.TextMatrix(i, 3), 60, 23)
     Desc = JustificaTextoCadenaOrdenCompra(fgDetalle.TextMatrix(I, 3), 40, 52)
  '***********************************************************

'    Desc = fgDetalle.TextMatrix(i, 2)
'        If Len(Desc) > 60 Then
'           div = Round(Len(Desc) / 60)
'            For j = 0 To div
'                If j = 0 Then
'                     DescF = Mid(Trim(Desc), 1, 60) & oImpresora.gPrnSaltoLinea & Space(12) & Space(20)  'pasar siguiente linea
'                End If
'
'                If j > 0 And j < div Then
'                      DescF = DescF & Mid(Trim(Desc), (60 * j) + 1, 60) & oImpresora.gPrnSaltoLinea & Space(12) & Space(20) 'pasar siguiente linea
'                End If
'
'                If j = div Then
'                   DescF = DescF & Mid(Trim(Desc), (60 * j) + 1, 60) & Space(60 - Len(Trim(Mid(Desc, (60 * j) + 1, 60)))) '& Space(4)
'                End If
'
'
'            Next
'            Desc = ImpreCarEsp(DescF)
'
'            If lbBienes Then
'
'               lsCadena = lsCadena & Space(13) & fgDetalle.TextMatrix(i, 4) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(i, 4)))) & Space(1) & fgDetalle.TextMatrix(i, 3) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(i, 3)))) & Space(2) & Desc & _
'                          JDER(Trim(fgDetalle.TextMatrix(i, 5)), 15) & JDER(Trim(fgDetalle.TextMatrix(i, 7)), 15) & oImpresora.gPrnSaltoLinea
'
'            Else
'               lsCadena = lsCadena & Space(13) & i & Space(18) & Desc & Space(1) & JDER(fgDetalle.TextMatrix(i, 7), 31) & oImpresora.gPrnSaltoLinea
'            End If
'        Else
'
'            If lbBienes Then
'               lsCadena = lsCadena & Space(13) & fgDetalle.TextMatrix(i, 4) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(i, 4)))) & Space(1) & fgDetalle.TextMatrix(i, 3) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(i, 3)))) & Space(2) & JIZQ(Desc, 60) & _
'                          JDER(Trim(fgDetalle.TextMatrix(i, 5)), 15) & JDER(Trim(fgDetalle.TextMatrix(i, 7)), 15) & oImpresora.gPrnSaltoLinea
'
'            Else
'               lsCadena = lsCadena & Space(13) & i & Space(18) & JIZQ(Desc, 60) & JDER(fgDetalle.TextMatrix(i, 7), 32) & oImpresora.gPrnSaltoLinea
'            End If
'
'
'        End If
        
           Dim a As Integer
           a = TextoFinLen
           If lbBienes Then
              'ALPA 20100409***********************************
              'lsCadena = lsCadena & Space(10) & Mid(Replace(Replace(GetAgencias(fgDetalle.TextMatrix(i, 1)), "Agencia", "Ag."), "Oficina", "Of."), 1, 20) & IIf(Len(Mid(Replace(GetAgencias(fgDetalle.TextMatrix(i, 1)), "Agencia", "Ag."), 1, 10)) < 10, Space(10 - Len(Mid(Replace(GetAgencias(fgDetalle.TextMatrix(i, 1)), "Agencia", "Ag."), 1, 10))), "") & Space(3) & fgDetalle.TextMatrix(i, 5) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(i, 5)))) & Space(1) & fgDetalle.TextMatrix(i, 4) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(i, 4)))) & Space(3) & Trim(Desc) & Space(40 - a) & _
              '            JDER(Trim(fgDetalle.TextMatrix(i, 6)), 14) & JDER(Trim(fgDetalle.TextMatrix(i, 8)), 15) & oImpresora.gPrnSaltoLinea
              lsCadena = lsCadena & Space(10) & Mid(Replace(Replace(GetAgencias(fgDetalle.TextMatrix(I, 1)), "Agencia", "Ag."), "Oficina", "Of."), 1, 20) & IIf(Len(Mid(Replace(GetAgencias(fgDetalle.TextMatrix(I, 1)), "Agencia", "Ag."), 1, 10)) < 10, Space(10 - Len(Mid(Replace(GetAgencias(fgDetalle.TextMatrix(I, 1)), "Agencia", "Ag."), 1, 10))), "") & Space(3) & fgDetalle.TextMatrix(I, 5) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(I, 5)))) & Space(1) & fgDetalle.TextMatrix(I, 4) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(I, 4)))) & Space(3) & Trim(Desc) & Space(40 - a) & _
                          JDER(Trim(fgDetalle.TextMatrix(I, 6)), 14) & JDER(Trim(fgDetalle.TextMatrix(I, 8)), 18) & oImpresora.gPrnSaltoLinea
              '************************************************

            Else
               If Len(Desc) > 60 Then
                    'ALPA 20100409*************************************
                    'lsCadena = lsCadena & Space(10) & Mid(Replace(Replace(GetAgencias(fgDetalle.TextMatrix(i, 1)), "Agencia", "Ag."), "Oficina", "Of."), 1, 20) & IIf(Len(Mid(Replace(GetAgencias(fgDetalle.TextMatrix(i, 1)), "Agencia", "Ag."), 1, 10)) < 10, Space(10 - Len(Mid(Replace(GetAgencias(fgDetalle.TextMatrix(i, 1)), "Agencia", "Ag."), 1, 10))), "") & Space(3) & fgDetalle.TextMatrix(i, 5) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(i, 5)))) & Space(1) & fgDetalle.TextMatrix(i, 4) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(i, 4)))) & Space(2) & Trim(Desc) & Space(40 - a) & _
                    '            JDER(Trim(fgDetalle.TextMatrix(i, 6)), 14) & JDER(Trim(fgDetalle.TextMatrix(i, 8)), 18) & oImpresora.gPrnSaltoLinea
                    lsCadena = lsCadena & Space(10) & Mid(Replace(Replace(GetAgencias(fgDetalle.TextMatrix(I, 1)), "Agencia", "Ag."), "Oficina", "Of."), 1, 20) & IIf(Len(Mid(Replace(GetAgencias(fgDetalle.TextMatrix(I, 1)), "Agencia", "Ag."), 1, 10)) < 10, Space(10 - Len(Mid(Replace(GetAgencias(fgDetalle.TextMatrix(I, 1)), "Agencia", "Ag."), 1, 10))), "") & Space(3) & fgDetalle.TextMatrix(I, 5) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(I, 5)))) & Space(1) & fgDetalle.TextMatrix(I, 4) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(I, 4)))) & Space(2) & Trim(Desc) & Space(40 - a) & _
                                JDER(Trim(fgDetalle.TextMatrix(I, 6)), 14) & JDER(Trim(fgDetalle.TextMatrix(I, 8)), 18) & oImpresora.gPrnSaltoLinea
                    '**************************************************
               Else
                    lsCadena = lsCadena & Space(13) & Space(19) & JIZQ(Desc, 60) & JDER(fgDetalle.TextMatrix(I, 8), 32) & oImpresora.gPrnSaltoLinea
               End If
            End If
       
    I = I + 1
    lineas = lineas + 1
  Loop
  '*****************************************
lineas = lineas + I

  lsCadena = lsCadena & Space(10) & String(115, "_") & oImpresora.gPrnSaltoLinea
  
  Dim lsMonto As String * 74
  lsMonto = Trim(CStr(ConvNumLet(nVal(txtTot.Text))))
  Dim nSubRenta As Currency
  Dim nRenta4 As Currency
  
  If lbBienes Then 'Orden Compra
  
  
  
    'nSubt = Format(((txtTot.Text) / 1.19), "0.00")
    nSubt = Format(((txtTot.Text) / (1 + gnIGV)), "0.00") ''*** PEAC 20110228
    
    nIGVT = txtTot.Text - nSubt

        If chkIGV.value = 1 Then
           '''lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(Format(nSubt, gcFormView), 18) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
           lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. ") & JDER(Format(nSubt, gcFormView), 18) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
           lsCadena = lsCadena & Space(15) & Space(74) & "I.G.V." & JDER(Format(nIGVT, gcFormView), 27) & oImpresora.gPrnSaltoLinea
           lineas = lineas + 2
        Else
           '''lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(Format(txtTot.Text, gcFormView), 18) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
           lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. ") & JDER(Format(txtTot.Text, gcFormView), 18) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
           lineas = lineas + 1
        End If
     
        lsCadena = lsCadena & Space(15) & Space(97) & String(11, "_") & oImpresora.gPrnSaltoLinea
        '''lsCadena = lsCadena & Space(10) & "SON: " & lsMonto & "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(txtTot.Text, 18) & oImpresora.gPrnSaltoLinea 'marg ers044-2016
        lsCadena = lsCadena & Space(10) & "SON: " & lsMonto & "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. ") & JDER(txtTot.Text, 18) & oImpresora.gPrnSaltoLinea 'marg ers044-2016
        lineas = lineas + 2
  Else
    If chkRenta4.value = 1 Then
       'nSubRenta = Format(((CCur(txtTot.Text)) / 8), "0.00") 'Comentado PASI20150211
       nSubRenta = Format(((CCur(txtTot.Text)) * 0.08), "0.00") 'PASI20151102 Cambio la Retencion de 10 a 8 segun TIC1502100009
       nRenta4 = txtTot.Text - nSubRenta
       
       '''lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(Format(txtTot.Text, gcFormView), 20) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
       lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. ") & JDER(Format(txtTot.Text, gcFormView), 20) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
       lsCadena = lsCadena & Space(15) & Space(74) & "Ret. 4ta." & JDER(Format(nSubRenta, gcFormView), 26) & oImpresora.gPrnSaltoLinea
       lsCadena = lsCadena & Space(15) & Space(98) & String(11, "_") & oImpresora.gPrnSaltoLinea
       '''lsCadena = lsCadena & Space(10) & "SON: " & lsMonto & "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(Format(nRenta4, gcFormView), 20) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
       lsCadena = lsCadena & Space(10) & "SON: " & lsMonto & "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. ") & JDER(Format(nRenta4, gcFormView), 20) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
       lineas = lineas + 4
     Else
         If chkIGV.value = 1 Then
            'nSubt = Format(((txtTot.Text) / 1.19), "0.00")
            nSubt = Format(((txtTot.Text) / (1 + gnIGV)), "0.00") '''*** PEAC 20110228
            
            nIGVT = txtTot.Text - nSubt
         
            '''lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(Format(nSubt, gcFormView), 20) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. ") & JDER(Format(nSubt, gcFormView), 20) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lsCadena = lsCadena & Space(15) & Space(74) & "I.G.V." & JDER(Format(nIGVT, gcFormView), 29) & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Space(15) & Space(98) & String(11, "_") & oImpresora.gPrnSaltoLinea
            '''lsCadena = lsCadena & Space(10) & "SON: " & lsMonto & "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(txtTot.Text, 20) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lsCadena = lsCadena & Space(10) & "SON: " & lsMonto & "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. ") & JDER(txtTot.Text, 20) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lineas = lineas + 4
         Else
            '''lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(Format(txtTot.Text, gcFormView), 20) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. ") & JDER(Format(txtTot.Text, gcFormView), 20) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lsCadena = lsCadena & Space(15) & Space(98) & String(11, "_") & oImpresora.gPrnSaltoLinea
            '''lsCadena = lsCadena & Space(10) & "SON: " & lsMonto & "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(txtTot.Text, 20) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lsCadena = lsCadena & Space(10) & "SON: " & lsMonto & "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, gcPEN_SIMBOLO, "$. ") & JDER(txtTot.Text, 20) & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
            lineas = lineas + 3
         End If
    End If
  End If

  lsCadena = lsCadena & Space(10) & String(115, "_") & oImpresora.gPrnSaltoLinea
  lsCadena = lsCadena & Space(10) & lsNegritaOn & "Concepto y Observaciones:" & lsNegritaOff & oImpresora.gPrnSaltoLinea
  lineas = lineas + 2
  
  sGlosa = Trim(txtMovDesc.Text)
  If Len(sGlosa) > 110 Then
     div = Round(Len(sGlosa) / 110)
       
     For J = 0 To div
         If J = 0 Then
            sGlosaF = Mid(sGlosa, 1, 110) & oImpresora.gPrnSaltoLinea & Space(10)
         End If
         If J = div Then
            sGlosaF = sGlosaF & Mid(sGlosa, (110 * J) + 1, 90)
         End If
         If J > 0 And J < div Then
            sGlosaF = sGlosaF & Mid(sGlosa, (110 * J) + 1, 110) & oImpresora.gPrnSaltoLinea & Space(38)
         End If
     Next
     sGlosa = sGlosaF
     
     lineas = lineas + div
  End If
  
  lineas = lineas + 1
  
  lsCadena = lsCadena & Space(10) & ImpreCarEsp(sGlosa) & oImpresora.gPrnSaltoLinea
  lsCadena = lsCadena & Space(10) & String(115, "_") & oImpresora.gPrnSaltoLinea
  
  
     lsCadena = lsCadena & Space(10) & " Los precios incluyen todos los tributos, seguros, transportes, inspecciones, pruebas y cualquier otro concepto" & oImpresora.gPrnSaltoLinea
     lsCadena = lsCadena & Space(10) & " que pueda incidir sobre el costo del(os) bien(es). Adjuntar copia de la orden de " & IIf(lbBienes, "compra ", "servicio ") & "a la factura." & oImpresora.gPrnSaltoLinea
     lsCadena = lsCadena & Space(20) & lsNegritaOn & " FACTURAR A NOMBRE: " & lsNegritaOff & "CMAC MAYNAS S.A.; RUC: 20103845328; Jr. Próspero N° 791 - Iquitos." & oImpresora.gPrnSaltoLinea
       
     'lsCadena = lsCadena & Space(10) & oImpresora.gPrnSaltoLinea
     lineas = lineas + 3
  
     lsCadena = lsCadena & Space(10) & lsNegritaOn & " PLAZO DE ENTREGA:" & lsNegritaOff & Space(3) & txtPlazo & Space(8) & lsNegritaOn & " FORMA DE PAGO:   " & lsNegritaOff & Space(3) & Trim(Left(cboFormaPago, Len(cboFormaPago) - 1)) & Space(8) & lsNegritaOn & "GARANTIA:" & lsNegritaOff & Space(3) & txtDiasAtraso.Text & Space(2) & "meses" & oImpresora.gPrnSaltoLinea '& oImpresora.gPrnSaltoLinea
     lineas = lineas + 1
     

       If lbBienes Then
          lsCadena = lsCadena & Space(10) & lsNegritaOn & "IMPORTANTE:" & lsNegritaOff & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "1. EL número de esta Orden debera consignarse en facturas, guías remisión, embalaje y correspondencia" & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "   respectiva. Asimismo el Contratista debe acusar recibo de esta Orden inmediatamente después de su recepción" & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "2. La CMAC MAYNAS S.A. se reserva el derecho de devolver los bienes, si no cumplen con las especificaciones" & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "   técnicas mínimas requeridas, y de acuerdo a lo ofertado por el Contratista, así como de anular la Orden de" & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "   Compra, según el procedimiento establecido en el reglamento de adquisiciones y contrataciones para el sistema CMAC." & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "3. El Contratista se compromete a cumplir las obligaciones que le corresponden, bajo sanción de quedar" & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "   inhabilitado para contratar con el Estado en caso de incumplimiento." & oImpresora.gPrnSaltoLinea
          '[LARI 20200331 ***********************************
          lsCadena = lsCadena & Space(10) & "4. En caso de retraso injustificado en la entrega de los bienes,  la CAJA aplicará al contratista " & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "   una penalidad por cada día de atraso, hasta por un monto máximo equivalente al Diez por ciento (10%) del monto de la ODC vigente." & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "   Cálculo de penalidad diaria=(0.10 x Monto)/(F x Plazo en días)" & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "   -Para plazos menores o iguales a sesenta (60) días, para bienes y servicios F=0.40" & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "   -Para plazos mayores a sesenta (60) días, para bienes y servicio F=0.25" & oImpresora.gPrnSaltoLinea
          
          lineas = lineas + 14
          Else
          lsCadena = lsCadena & Space(10) & "1. En caso de retraso injustificado en la entrega de los bienes,  la CAJA aplicará al contratista " & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "   una penalidad por cada día de atraso, hasta por un monto máximo equivalente al Diez por ciento (10%) del monto de la ODC vigente." & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "   Cálculo de penalidad diaria=(0.10 x Monto)/(F x Plazo en días)" & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "   -Para plazos menores o iguales a sesenta (60) días, para bienes y servicios F=0.40" & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "   -Para plazos mayores a sesenta (60) días, para bienes y servicio F=0.25" & oImpresora.gPrnSaltoLinea
          lineas = lineas + 6
          '**************************************************]
       End If
       
       For k = lineas To 40
           lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
       Next
    
       lsCadena = lsCadena & Space(10) & String(115, "_") & oImpresora.gPrnSaltoLinea
       'ALPA 20090730***********************************************************
       'lsCadena = lsCadena & Space(14) & "LOGISTICA" & Space(35) & "GERENCIA" & Space(45) & "PROVEEDOR" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       lsCadena = lsCadena & Space(14) & "AREA USUARIA" & Space(17) & "LOGISTICA" & Space(20) & "GERENCIA" & Space(30) & "PROVEEDOR" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
       '************************************************************************
       lsCadena = lsCadena & Space(10) & String(115, "_") & oImpresora.gPrnSaltoLinea
       lsCadena = lsCadena & Space(10) & Space(10) & "Cualquier consulta: Teléfono (065) 58-1770 Anexos 1412 - 1413 / Fax Anexo 1414" & oImpresora.gPrnSaltoLinea
       'lsCadena = lsCadena & Space(10) & Space(15) & "Correos: cpanduro@cmacmaynas.com.pe,  gurteaga@cmacmaynas.com.pe, dleveau@cmacmaynas.com.pe" & oImpresora.gPrnSaltoLinea
       lsCadena = lsCadena & Space(10) & "Correos: " & lsCorreos & oImpresora.gPrnSaltoLinea 'EJVG20131126
       lsCadena = lsCadena & Space(10) & String(115, "_") & oImpresora.gPrnSaltoLinea

  

oPrevio.Show lsCadena, Caption, , , gIBM

        'ARLO 20170125
        If (cmbmoneda.ListIndex = 1) And sDocDesc = "Orden de Compra" Then
        gsOpeCod = LogPistasImpresionOrdenComprasSoles
        ElseIf (cmbmoneda.ListIndex <> 1) And sDocDesc = "Orden de Compra" Then
        gsOpeCod = LogPistasImpresionOrdenComprasDolares
        ElseIf (cmbmoneda.ListIndex = 1) And sDocDesc = "Orden de Servicio" Then
        gsOpeCod = LogPistasImpresionOrdenServicioSoles
        Else: gsOpeCod = LogPistasImpresionOrdenServicioDolares
        End If
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Imprimio la " & sDocDesc & " " & sPreNombre & " en " & Mid(cmbmoneda.Text, 1, 7) & " N°: " & gcDocNro, 3
        Set objPista = Nothing
        '***********
End Sub



'Private Sub cmdDocumento_Click()
'Dim N As Integer
'Dim sTit   As String
'Dim sCabe  As String
'Dim sTexto As String
'Dim sPie   As String
'Dim sDet   As String
'Dim sDesc  As String
'Dim nLin   As Integer
'Dim nLinD  As Integer
'Dim P      As Integer
'Dim nTabulador As Integer
'
'Dim sPrenombre As String
'
'Dim oCon   As DConecta
'Set oCon = New DConecta
'oCon.AbreConexion
'P = 0
'
'CON = PrnSet("C+")
'COFF = PrnSet("C-")
'BON = PrnSet("B+")
'BOFF = PrnSet("B-")
'
'If Right(cmbTipoOC.Text, 1) = "D" Then
'    sPrenombre = " DIRECTA "
'ElseIf Right(cmbTipoOC.Text, 1) = "P" Then
'    sPrenombre = " PROCESO "
'End If
'
'If lbBienes Then
'   If Right(cmbTipoOC.Text, 1) = "D" Or Right(cmbTipoOC.Text, 1) = "P" Then
'        sTit = "ORDEN DE COMPRA Nro. " & gcDocNro & " ---- " & " O/C" & sPrenombre & " (" & Right(lblDocOCD.Caption, 13) & ")"
'   Else
'        sTit = "O R D E N   D E   C O M P R A   Nro. " & gcDocNro
'   End If
'Else
'   sTit = "ORDEN DE SERVICIO Nro. " & gcDocNro & " ---- " & " O/S" & sPrenombre & "( " & Right(lblDocOCD.Caption, 13) & " ) "
'End If
'gdFecha = txtFecha
'Lin1 sTexto, CabeOrden(P, nLin, sTit), 0
'If chkProy.value = vbChecked Then
'   Lin1 sTexto, "Proyecto     : " & BON & "[" & txtProy.Tag & "] " & txtProy & BOFF
'   nLin = nLin + 1
'End If
'Lin1 sTexto, PrnSet("I+") & "Por medio de la presente solicitamos se sirvan efectuar lo siguiente:" & PrnSet("I-"), 0
'nLin = nLin + 1
'
'sDet = ""
'For N = 1 To fgDetalle.Rows - 1
'  If fgDetalle.TextMatrix(N, 1) <> "" Then
'    If nVal(fgDetalle.TextMatrix(N, 7)) <> 0 Then
'      nLin = nLin + 1
'      gsGlosa = Replace(fgDetalle.TextMatrix(N, 2), Chr(13), " ")
'      gsGlosa = Replace(gsGlosa, oImpresora.gPrnSaltoLinea, " ")
'
'      Lin1 sDet, CON & ImpreGlosa(IIf(fgDetalle.TextMatrix(N, 1) <> "", Format(Left(fgDetalle.TextMatrix(N, 0), 25), "00"), "  ") & _
'                 " " & Left(fgDetalle.TextMatrix(N, 1) + Space(13), 18) & _
'                 " ", 22 + 75, False, nLin), 0
'
'      sDet = Left(sDet, Len(sDet) - 3)
'
'      Lin1 sDet, " " & Left(fgDetalle.TextMatrix(N, 3) + Space(5), 7), 0
'      Lin1 sDet, " " & Right(Space(5) & fgDetalle.TextMatrix(N, 4), 10), 0
'
'      If lbBienes = True Then
'            Lin1 sDet, " " & Right(Space(8) & Format(fgDetalle.TextMatrix(N, 7) / fgDetalle.TextMatrix(N, 4), "#######.#0"), 10), 0
'         Else
'            Lin1 sDet, " " & "", 0
'      End If
'      If lbBienes = True Then
'            Lin1 sDet, " " & Right(Space(10) & fgDetalle.TextMatrix(N, 7), 10) & COFF
'         Else
'            If Len(sDet) > 100 Then
'               nTabulador = 100 - (Len(sDet) Mod 100)
'            End If
'            If (Len(sDet) Mod 100) > 8 And (Len(sDet) Mod 100) < 40 Then
'                sDet = sDet + String(nTabulador, " ")
'               Else
'            End If
'
'            Lin1 sDet, " " & Space(16) & Right(Space(10) & fgDetalle.TextMatrix(N, 7), 10) & COFF
'      End If
'
'
'      If nLin > gnLinPage - 12 Then
'         Lin1 sDet, CabeOrden(P, nLin, sTit) & "<<SALTO>>", 0
'         nLin = nLin + 3
'      End If
'    End If
'  End If
'Next
'If P = 1 And nLin < 28 Then
'   Lin1 sDet, "", 28 - nLin
'   nLin = 28
'End If
''******************************* Impresion Cabecera ************************
'Lin1 sTexto, ImpreDetLog(sDet, "Unidad    Cantidad    Pre.Unit", "SubTotal", 20 - N), 0
'nLin = nLin + 3
'Lin1 sTexto, CON & BON & Justifica(" SON : " & ConvNumLet(nVal(txtTot)), 110) & " ", 0
'Lin1 sTexto, "TOTAL " & Right(Space(14) & gsSimbolo & " " & txtTot, 14) & BOFF & COFF
'Lin1 sTexto, CON & Space(109) & String(25, "-") & COFF
'nLin = nLin + 2
'If nLin > gnLinPage - 21 Then
'   Lin1 sTexto, CabeOrden(P, nLin, sTit), 0
'End If
'gsGlosa = txtMovDesc
'Lin1 sTexto, BON & ImpreGlosa("Observaciones : ", gnColPage, , nLin) & BOFF
'
'Dim sMsg As String
''sMsg = "1. Estos Costos Incluyen los Impuestos de Ley " & oImpresora.gPrnSaltoLinea
'sMsg = "1. Facturar a nombre de la " & Trim(gsEmpresaCompleto) & " " & Trim(gsEmpresaDireccion) & oImpresora.gPrnSaltoLinea & "   R.U.C. " & gsRUC & oImpresora.gPrnSaltoLinea _
'     & "2. Consignar el número de la presente " & IIf(lbBienes, "Orden de Compra", "Orden de Servicio") & " en su Factura y Guía de Remisión." '& oImpresora.gPrnSaltoLinea _
''    & "4. La Factura deberá ser enviada inmediatamente a la Caja Municipal al recibir este documento"
'Lin1 sTexto, CON & BON & Centra("***********   IMPORTANTE  **************", gnColPage) & BOFF
'Lin1 sTexto, sMsg & COFF, 2
'nLin = nLin + 6
'
'Lin1 sTexto, BON & "CONDICIONES :" & BOFF
'Lin1 sTexto, "Forma de Pago    : " & Justifica(Trim(Left(cboFormaPago, Len(cboFormaPago) - 1)), 20) & "  Plazo de Entrega : " & Justifica(txtPlazo, 10)
'Lin1 sTexto, "Lugar de Entrega : " & txtLugarEntrega
'nLin = nLin + 3
''Lin1 sTexto, ImprimePiePagina("248")
'Lin1 sTexto, ImprimePiePagina("248")
'
'sSQL = "SELECT m.cMovNro, mr.nMovNro, m.nMovEstado, m.nMovFlag FROM MovRef mr JOIN Mov m ON m.nMovNro = mr.nMovNroRef WHERE nMovNroRef = '" & gcMovNro & "'"
'Set rs = oCon.CargaRecordSet(sSQL)
'If Not rs.EOF Then
'   If rs!nMovEstado = 16 And rs!nMovFlag = gMovFlagVigente Then
'      Lin1 sTexto, BON & Centra("*** A P R O B A D O   P O R   P R E S U P U E S T O *** " & GetFechaMov(rs!cMovNro, True) & " " & Mid(rs!cMovNro, 9, 6), 79) & BOFF
'   ElseIf rs!nMovEstado = 14 Then
'      Lin1 sTexto, BON & Centra("*** R E C H A Z A D O   P O R   P R E S U P U E S T O *** " & GetFechaMov(rs!cMovNro, True) & " " & Mid(rs!cMovNro, 9, 6), 79) & BOFF
'   Else
'      Lin1 sTexto, BON & Centra("*** ESTADO DE DOCUMENTO : " & UCase(lsEstado) & " ***", 79) & BOFF
'   End If
'ElseIf UCase(lsEstado) = "APROBADO" Then
'   Lin1 sTexto, BON & Centra("*** A P R O B A D O   P O R   P R E S U P U E S T O *** ", 79) & BOFF
'Else
'   Lin1 sTexto, BON & Centra("*** ESTADO DE DOCUMENTO : " & UCase(lsEstado) & " ***", 79) & BOFF
'End If
'nLin = nLin + 9
'Lin1 sTexto, "", gnLinPage - 5 - nLin
'sTexto = ImpreCarEsp(sTexto)
'frmImpreSeleCopia.Inicio sTexto, sDocDesc
''frmPrevio.Previo rtxtAsiento, "Documento: " & lblDoc.Caption, False, gnLinPage
'End Sub



Private Function CabeOrden(ByRef P As Integer, ByRef nLin As Integer, sTit As String) As String
Dim sCabe   As String
Dim ldFecha As Date
If P > 0 Then sCabe = oImpresora.gPrnSaltoPagina
P = P + 1
ldFecha = gdFecSis
gdFecSis = txtFecha
'''Lin1 sCabe, CabeRepoAnt("", "", 80, Trim(gsNomAge), "", sTit, IIf(gsSimbolo = gcMN, "M.N.", "M.E."), 1), 2 'MARG ERS044-2016
Lin1 sCabe, CabeRepoAnt("", "", 80, Trim(gsNomAge), "", sTit, IIf(gsSimbolo = gcPEN_SIMBOLO, "M.N.", "M.E."), 1), 2 'MARG ERS044-2016
gdFecSis = ldFecha
Lin1 sCabe, Space(47) & "Ref. Cotización Nro " & txtCotNro
Lin1 sCabe, "Razon Social: " & BON & Justifica(txtProvNom, 46) & BOFF & " R.U.C.: " & Me.txtProvRuc
Lin1 sCabe, "Dirección : " & Justifica(txtProvDir, 48) & " TELF. : " & txtProvTele
Lin1 sCabe, "Area Usuaria : " & BON & "[" & txtArea.Text & "] " & txtAgeDesc & BOFF
nLin = 10
CabeOrden = sCabe
End Function

Private Sub cmdExaminar_Click()
If fgDetalle.TextMatrix(fgDetalle.row, 11) = "B" Then
   'ExaminarObjeto
   ExaminarObjeto fsObjetoCod 'EJVG20140319
Else
   ExaminarServicio
End If
If fgDetalle.TextMatrix(fgDetalle.row, 2) <> "" And fgDetalle.TextMatrix(fgDetalle.row, 12) = "" Then
   fgDetalle.TextMatrix(fgDetalle.row, 12) = txtPlazo
End If
End Sub

Private Sub ExaminarServicio()
    Dim sSqlO As String
    Dim rsObj As ADODB.Recordset
    Set rsObj = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim nNiv  As Integer
    Dim sCtaCod  As String
    Dim sCtaDes  As String
    
    oCon.AbreConexion
    
    sSqlO = "SELECT DISTINCT a.cCtaContCod as cObjetoCod, b.cCtaContDesc, 2 as nObjetoNiv " _
          & "FROM  " & gcCentralCom & "OpeCta a,  " & gcCentralCom & "CtaCont b " _
          & "WHERE b.cCtaContCod = a.cCtaContCod AND ((a.cOpeCod='" & gcOpeCod & "') AND (a.cOpeCtaDH='D'))"
    
    Set rs = oCon.CargaRecordSet(sSqlO)
    If rs.EOF Then
        MsgBox "No se asignaron Conceptos a Operación", vbCritical, "Error"
        txtObj.SetFocus
        Exit Sub
    End If
    frmDescObjeto.Inicio rs, txtObj.Text, 2, "Servicios"
    
    If frmDescObjeto.lOk Then
       sCtaCod = gaObj(0, 0, 0): sCtaDes = gaObj(0, 1, 0)
       AsignaObjetosSer sCtaCod, sCtaDes
       fgDetalle.Enabled = True
       txtObj.Visible = False
       cmdExaminar.Visible = False
       fgDetalle.col = 8
       fgDetalle.SetFocus
    Else
       If txtObj.Visible And txtObj.Enabled Then txtObj.SetFocus
    End If
    If Not rs Is Nothing Then
       If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
    End If
    If fgDetalle.TextMatrix(fgDetalle.row, 0) <> "" Then
       RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.row, 0)
    End If
End Sub

Private Sub ExaminarObjeto(Optional ByVal psBSCod As String = "")
If fgDetalle.TextMatrix(fgDetalle.row, 1) <> "" Then
    Dim sSqlO As String
    Dim lsCodigo As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oAlmacen As New DLogAlmacen
    'EJVG20140319 ***
    If Len(psBSCod) = 0 Then
        Set rs = oAlmacen.GetBienesAlmacen(, "11','12','13")
        If rs.EOF Then
            MsgBox "No existen Bienes", vbCritical, "Error"
            If txtObj.Visible And txtObj.Enabled Then txtObj.SetFocus
            Exit Sub
        End If
    End If
    Set oAlmacen = Nothing
    Dim oDesc As New ClassDescObjeto
    oDesc.lbUltNivel = True
    If Len(psBSCod) = 0 Then
        oDesc.Show rs, txtObj
    End If
    'END EJVG *******
    'If oDesc.lbOk Then
    If (Len(psBSCod) = 0 And oDesc.lbOK) Or (Len(psBSCod) > 0) Then 'EJVG20140319
       'lsCodigo = oDesc.gsSelecCod
       lsCodigo = IIf(Len(psBSCod) = 0, oDesc.gsSelecCod, psBSCod) 'EJVG20140319
       oCon.AbreConexion
       'ALPA 20091010************************************************************
       'sSqlO = "Select cSubCtaCod from areaagencia where cAreaCod + cAgeCod = '" & Me.txtArea.Text & "'"
       sSqlO = "Select cSubCtaCod from areaagencia where cAgeCod = '" & fgDetalle.TextMatrix(fgDetalle.row, 1) & "'"
       '*************************************************************************
       Set rs = oCon.CargaRecordSet(sSqlO)
       If rs.BOF = True Or rs.EOF = True Then
            MsgBox "Agencia no tiene subcuenta en este bien/servicio"
            Exit Sub
       End If
       sSqlO = FormaSelect(gcOpeCod, lsCodigo, 0, rs.Fields(0))

       Set rs = oCon.CargaRecordSet(sSqlO)
       
       If RSVacio(rs) Then
          MsgBox "Objeto no asignado a Operación", vbInformation, "¡Aviso!"
          If txtObj.Visible And txtObj.Enabled Then txtObj.SetFocus
          Exit Sub
       End If
       sObjCod = Trim(rs!cObjCod)
       sObjDesc = Trim(rs!cObjDesc)
       sObjUnid = IIf(IsNull(rs!cConsDescripcion) Or Trim(rs!cConsDescripcion) = "", "UND", rs!cConsDescripcion)
       If rs.RecordCount = 1 Then
          sCtaCod = Trim(rs!cObjetoCod)
          sCtaDesc = Trim(rs!cObjDesc)
       Else
          frmDescObjeto.Inicio rs, "", 1, "Cuentas Contables"
          If frmDescObjeto.lOk Then
             sCtaCod = gaObj(0, 0, UBound(gaObj, 3) - 1)
             sCtaDesc = gaObj(0, 1, UBound(gaObj, 3) - 1)
          Else
             MsgBox "No se definió Cuenta Contable", vbInformation, "¡Aviso!"
             If txtObj.Enabled And txtObj.Visible Then txtObj.SetFocus
             Exit Sub
          End If
       End If
       ActualizaFG fgDetalle.row
       fgDetalle.Enabled = True
       txtObj.Visible = False
       cmdExaminar.Visible = False
       If fgDetalle.Visible And fgDetalle.Enabled Then fgDetalle.SetFocus
       rs.Close
    Else
       If txtObj.Enabled And txtObj.Visible Then
          txtObj.SetFocus
       End If
    End If
    Else
    MsgBox "Seleccionar Agencia Destino", vbApplicationModal
    End If
End Sub

Private Sub fgDetalle_DblClick()
If lbImprime Or lbRegComp Then 'WIOR 20130110 AGREGO lbRegComp
   Exit Sub
End If
'If lbModifica And (gcOpeCod = "501208" Or gcOpeCod = "502208") And fgDetalle.col = 2 Then 'EJVG20140225 , PASI20150917 ERS0472015 comento el codigo
'    Exit Sub
'End If
If fgDetalle.col = 2 Then
   EnfocaTexto txtObj, 0, fgDetalle
End If
If (fgDetalle.col = 5 And fgDetalle.TextMatrix(fgDetalle.row, 11) = "B") Or (fgDetalle.col = 6 And fgDetalle.TextMatrix(fgDetalle.row, 11) = "B") Or fgDetalle.col = 8 Then
   EnfocaTexto txtCant, 0, fgDetalle
End If
If fgDetalle.col = 3 Then
   EnfocaTexto txtConcepto, 0, fgDetalle
End If
If fgDetalle.col = 12 Then
   EnfocaTexto txtFecPlazo, 0, fgDetalle
End If
If fgDetalle.col = 1 Then
   EnfocaTexto textObjDes, 0, fgDetalle
End If
End Sub

Private Sub fgDetalle_KeyPress(KeyAscii As Integer)
If lbImprime Or lbRegComp Then 'WIOR 20130110 AGREGO lbRegComp
   Exit Sub
End If
'If lbModifica And (gcOpeCod = "501208" Or gcOpeCod = "502208") And fgDetalle.col = 2 Then 'EJVG20140225, PASI20150917 ERS0472015 comento el codigo
'    Exit Sub
'End If

If fgDetalle.col = 2 Then
   If KeyAscii <> 32 Then
      EnfocaTexto txtObj, IIf(KeyAscii = 13, 0, KeyAscii), fgDetalle
   End If
End If
If (fgDetalle.col = 5 And fgDetalle.TextMatrix(fgDetalle.row, 11) = "B") _
Or (fgDetalle.col = 6 And fgDetalle.TextMatrix(fgDetalle.row, 11) = "B") _
Or (fgDetalle.col = 8) Then
   If InStr("0123456789.", Chr(KeyAscii)) > 0 Then
      EnfocaTexto txtCant, KeyAscii, fgDetalle
   Else
      If KeyAscii = 13 Then EnfocaTexto txtCant, 0, fgDetalle
   End If
End If
If fgDetalle.col = 3 Then 'And fgDetalle.TextMatrix(fgDetalle.Row, 10) = "S") Then
   EnfocaTexto txtConcepto, 0, fgDetalle
End If
If fgDetalle.col = 12 Then
   EnfocaTexto txtFecPlazo, 0, fgDetalle, , , "F"
End If
End Sub
Private Sub fgDetalle_KeyUp(KeyCode As Integer, Shift As Integer)
Dim k As Integer
If lbImprime Or lbRegComp Then 'WIOR 20130110 AGREGO lbRegComp
   Exit Sub
End If
'If lbModifica And (gcOpeCod = "501208" Or gcOpeCod = "502208") And fgDetalle.col = 2 Then 'EJVG20140225, PASI20150917 ERS0472015 comento el codigo
'    Exit Sub
'End If

If KeyCode = 46 Then
   If fgDetalle.col = 2 Or fgDetalle.col = 5 Or fgDetalle.col = 6 Or fgDetalle.col = 8 Then
      KeyUp_Flex fgDetalle, KeyCode, Shift
      If fgDetalle.col = 6 Then
         fgDetalle.TextMatrix(fgDetalle.row, 6) = ""
      End If
      If fgDetalle.col = 2 Then
         For k = 1 To fgDetalle.Cols - 1
            fgDetalle.TextMatrix(fgDetalle.row, k) = ""
         Next
      Else
         fgDetalle.TextMatrix(fgDetalle.row, 8) = ""
      End If
      SumasDoc
   End If
Else
   KeyUp_Flex fgDetalle, KeyCode, Shift
End If
End Sub

Private Sub Form_Activate()
    Dim ofun As NContFunciones
    Set ofun = New NContFunciones
    
    GetTipCambio gdFecSis, Not gbBitCentral
    
    If Not Trim(txtProy.Tag) = "" Then
       chkProy.value = vbChecked
    End If
    If lSalir Then
       Unload Me
    End If
    
    'If Me.cmbMoneda.Text = "" Then
        If Mid(gcOpeCod, 3, 1) = gcMNDig Then
            cmbmoneda.ListIndex = 1
        Else
            cmbmoneda.ListIndex = 0
        End If
    'Else
    '    cmbMoneda.ListIndex = 1
    'End If
    If Me.txtProvCod.Text = "" And Not lbModifica Then
        gcDocNro = ""
    End If
    lTransActiva = False
       
    
    If gcDocNro = "" And gcDocTpo <> "" Then
    
        If Me.cmbmoneda.Text <> "" Then gcDocNro = ofun.GeneraDocNro(CInt(gcDocTpo), Right(Me.cmbmoneda.Text, 1), Year(gdFecSis), "OSC") 'NAGL 20191212 Agregó "OSC"
        LblDoc.Caption = UCase(sDocDesc) & "   Nº " & gcDocNro
        
    End If
    
End Sub

Private Sub Form_Load()
    Dim n As Integer, nSaldo As Currency, k As Currency
    Dim sCtaCod As String, nItem As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oConst As DConstantes
    Set oConst = New DConstantes
    Dim oArea As DActualizaDatosArea
    Set oArea = New DActualizaDatosArea
    Dim oPer As UPersona
    Set oPer = New UPersona
    Dim sTipodoc  As String
    'WIOR 20130110 *****************************
    Dim oDoc As DOperacion
    Dim rsDoc As New ADODB.Recordset
    'WIOR FIN **********************************
    Dim ofun As NContFunciones
    Set ofun = New NContFunciones
    Dim lsNContrato As String 'PASI20140724 TI-ERS077-2014
    Dim olog As DLogGeneral 'PASI20140724 TI-ERS077-2014
    Set olog = New DLogGeneral
    Dim rsComboSum As ADODB.Recordset 'PASI20140724 TI-ERS077-2014
    Set rsComboSum = New ADODB.Recordset
    
   
    lSalir = False
    CargaCombo oConst.GetConstante(gMoneda), Me.cmbmoneda
    
    lOrdenCompra = True
    CargaComboSuministro 'PASIERS0772014
    
    'gnIGV
    
    If gcOpeCod = "501207" Or gcOpeCod = "502207" Then
        cmdSeleccion.Visible = False
    End If
    
    If Mid(gcOpeCod, 3, 1) = gcMEDig Then  'Identificación de Tipo de Moneda
       gsSimbolo = gcME
       If gnTipCambio = 0 Then
            GetTipCambio gdFecSis, Not gbBitCentral
       End If
       fraCambio.Visible = True
       txtTipCambio = Format(gnTipCambio, "##,###,##0.0000")
       
       If gbBitTCPonderado Then
            txtTipCompra = Format(gnTipCambioPonderado, "##,###,##0.0000")
            Label5.Caption = "Ponde."
       Else
            txtTipCompra = Format(gnTipCambioC, "##,###,##0.0000")
       End If
    Else
       '''gsSimbolo = gcMN 'MARG ERS044-2016
       gsSimbolo = gcPEN_SIMBOLO 'MARG ERS044-2016
    End If
    
    Set rs = CargaOpeCta(gcOpeCod, "H")
    
    If rs.EOF Then
       MsgBox "No se asignó Cuenta de Provisión de Bienes a Operación", vbInformation, "¡Aviso!"
       rs.Close
       Set rs = Nothing
       lSalir = True
       Exit Sub
    End If
    sCtaProvis = rs!cCtaContCod
    
    txtFecha = Format(gdFecSis, gsFormatoFechaView)
    
'    If lbModifica = True Then
'       txtDiasAtraso.Visible = True
'       lblDiasAtraso.Visible = True
'    End If
    
    Me.Caption = gcOpeDesc
    FormatoOrden
    FormatoObjeto
    EnumeraItems fgDetalle
    lnColorBien = "&H00F0FFFF"
    lnColorServ = "&H00FFFFC0"
    If lbBienes Then
       txtObj.BackColor = lnColorBien
       txtCant.BackColor = lnColorBien
       fgDetalle.BackColor = lnColorBien
       fgDetalle.BackColorBkg = lnColorBien
       lbltipooc.Caption = "Tipo O/C"
       
       If Mid(gcOpeCod, 3, 1) = 2 Then
          LblDoc.BackColor = &H80FF80
          txtTot.BackColor = &H80FF80
       End If
       
    Else
       txtObj.BackColor = lnColorServ
       txtCant.BackColor = lnColorServ
       fgDetalle.BackColor = lnColorServ
       fgDetalle.BackColorBkg = lnColorServ
       lbltipooc.Caption = "Tipo O/S"
       
       If Mid(gcOpeCod, 3, 1) = 2 Then
          LblDoc.BackColor = &H80FF80
          txtTot.BackColor = &H80FF80
       End If
       
       
    End If
    fgObj.BackColor = lnColorServ
    fgObj.BackColorBkg = lnColorServ
    
    If lbBienes Then
       sDocDesc = "Orden de Compra"
    Else
       sDocDesc = "Orden de Servicio"
    End If
    
    'EJVG20111115***************************
    For n = 1 To fgDetalle.Rows - 1
        If lbBienes Then
           fgDetalle.TextMatrix(n, 11) = "B"
        Else
           fgDetalle.TextMatrix(n, 11) = "S"
        End If
    Next
    '***************************************
    oCon.AbreConexion
    
    
    If sProvCod <> "" Then
       txtProvCod.Text = sProvCod
       
       Me.txtPersona.Text = sProvCod
       
       oPer.ObtieneClientexCodigo lsCodPers
       
       If oPer Is Nothing Then
          MsgBox "Proveedor no registrado. Por favor verificar", vbCritical, "Error"
          lSalir = True
          Exit Sub
       End If
       
       txtProvNom.Text = oPer.sPersNombre
       txtProvDir.Text = oPer.sPersDireccDomicilio
       txtProvTele.Text = oPer.sPersTelefono
       txtProvRuc.Text = oPer.sPersIdnroRUC
       txtFecha.Text = Format(gdFecha, "dd/mm/yyyy")
       txtMovDesc.Text = gsGlosa
       lsMovNro = gcMovNro
       lsMovNroEdit = gcMovNro 'PASI20151210
              
       sSql = "SELECT m.cAreaCod,ctipoOc, ISNULL(o.cAreaDescripcion,Age.cAgeDescripcion) cAreaDes , cCotizacNro, dMovPlazo, cMovFormaPago, cMovLugarEntrega, cPresuCod,isnull(m.nNroDiasAtraso,0) nNroDiasAtraso FROM MovCotizac m LEFT JOIN Areas O ON o.cAreaCod = m.cAreaCod LEFT JOIN Agencias age ON age.cAgeCod = RIGHT(m.cAreaCod,2) WHERE nmovnro = '" & gcMovNro & "'"
       'Exit Sub
       
       Set rs = oCon.CargaRecordSet(sSql)
       If Not rs.EOF Then
          txtArea.Text = rs!cAreaCod
          If Not IsNull(rs!cAreaDes) Then
            txtAgeDesc.Text = rs!cAreaDes
          End If
       
          txtCotNro.Text = rs!cCotizacNro
          txtPlazo.value = Format(rs!dMovPlazo, gsFormatoFechaView)
          Me.txtProy.Tag = rs!cPresuCod
          Me.txtLugarEntrega.Text = rs!cMovLugarEntrega
          Me.txtDiasAtraso.Text = rs!nNroDiasAtraso
       
          For nItem = 0 To 3
             If Right(cboFormaPago.List(nItem), 1) = rs!cMovFormaPago Then
                cboFormaPago.ListIndex = nItem
                Exit For
             End If
          Next
       
        If rs!ctipoOC = "D" Then
                lblDocOCD.Visible = True
                cmbTipoOC.ListIndex = 0 'Directa
            ElseIf rs!ctipoOC = "P" Then
                lblDocOCD.Visible = False
                cmbTipoOC.ListIndex = 1 'proceso
        End If
       'PASI20140819 TI-ERS077-2014
        lsNContrato = olog.ExisteMovContrato(gcMovNro)
        If lsNContrato <> "" Then
            Me.chkContSuministro.value = 1
            Me.cboContSuministro.ListIndex = IndiceListaCombo(cboContSuministro, lsNContrato)
        End If
        Me.chkContSuministro.Enabled = False
        Me.cboContSuministro.Enabled = False
        'end PASI
       
        cmbTipoOC.Enabled = False
    
          sSql = "SELECT nPresuCod cPresu, cPresuDescripcion cDesPre FROM PresuClase Where nPresuTpo = '" & txtProy.Tag & "'"
          Set rs = oCon.CargaRecordSet(sSql)
          If Not rs.EOF Then
             txtProy = rs!cDesPre
          End If
       End If
         Dim sBaseNew As String
         Dim rsObj    As ADODB.Recordset
         Dim sObjCod  As String
         sBaseNew = ""
         
         'marg ers044-2016
'         sSQL = " SELECT  mc.nMovNro, mc.nMovItem, mo.nMovObjOrden cMovObjOrden, mc.cCtaContCod, " _
'              & " mo.cObjetoCod, mcd.cDescrip, dItemPlazo, moc.nMovCant," & IIf(gsSimbolo = gcMN, "mc.nMovImporte ", "me.nMovMeImporte ") & " nMovImporte,cAgeCod " _
'              & " FROM " & sBaseNew & "MovCta mc " & IIf(gsSimbolo = gcMN, "", " JOIN " & sBaseNew & "MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem") _
'              & " LEFT JOIN ( SELECT nMovNro,nMovItem, nMovObjOrden, cObjetoCod FROM " & sBaseNew & "MovObj " _
'              & "             WHERE nMovNro = '" & gcMovNro & "') mo on mo.nMovNro = mc.nMovNro and mo.nMovItem = mc.nMovItem" _
'              & " LEFT JOIN " & sBaseNew & "MovCant moc ON moc.nMovNro = mo.nMovNro and moc.nMovItem = mo.nMovItem" _
'              & " LEFT JOIN " & sBaseNew & "MovCotizacDet mcd ON mcd.nMovNro = mc.nMovNro and mcd.nMovItem = mc.nMovItem " _
'              & " LEFT JOIN " & sBaseNew & "MovAgencia MA ON MA.nMovNro = mc.nMovNro and MA.nMovItem = mc.nMovItem " _
'              & " WHERE mc.nMovNro = '" & gcMovNro & "' and mc.nMovImporte <> 0 And mc.cCtaContCod Not Like  '25%' ORDER BY mc.nMovItem " 'MARG ERS044-2016
         sSql = " SELECT  mc.nMovNro, mc.nMovItem, mo.nMovObjOrden cMovObjOrden, mc.cCtaContCod, " _
              & " mo.cObjetoCod, mcd.cDescrip, dItemPlazo, moc.nMovCant," & IIf(gsSimbolo = gcPEN_SIMBOLO, "mc.nMovImporte ", "me.nMovMeImporte ") & " nMovImporte,cAgeCod " _
              & " FROM " & sBaseNew & "MovCta mc " & IIf(gsSimbolo = gcPEN_SIMBOLO, "", " JOIN " & sBaseNew & "MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem") _
              & " LEFT JOIN ( SELECT nMovNro,nMovItem, nMovObjOrden, cObjetoCod FROM " & sBaseNew & "MovObj " _
              & "             WHERE nMovNro = '" & gcMovNro & "') mo on mo.nMovNro = mc.nMovNro and mo.nMovItem = mc.nMovItem" _
              & " LEFT JOIN " & sBaseNew & "MovCant moc ON moc.nMovNro = mo.nMovNro and moc.nMovItem = mo.nMovItem" _
              & " LEFT JOIN " & sBaseNew & "MovCotizacDet mcd ON mcd.nMovNro = mc.nMovNro and mcd.nMovItem = mc.nMovItem " _
              & " LEFT JOIN " & sBaseNew & "MovAgencia MA ON MA.nMovNro = mc.nMovNro and MA.nMovItem = mc.nMovItem " _
              & " WHERE mc.nMovNro = '" & gcMovNro & "' and mc.nMovImporte <> 0 And mc.cCtaContCod Not Like  '25%' ORDER BY mc.nMovItem " 'MARG ERS044-2016
        'end marg
         'EJVG20140319 Se agregó MovAgencia
         Set rs = oCon.CargaRecordSet(sSql)
         n = 0
         nItem = 0
         Do While Not rs.EOF
            If nItem <> rs!nMovItem Then
               n = n + 1
               If n > fgDetalle.Rows - 1 Then
                  AdicionaRow fgDetalle
               End If
               fgDetalle.TextMatrix(n, 0) = n
               fgDetalle.TextMatrix(n, 2) = rs!cCtaContCod
               fgDetalle.TextMatrix(n, 8) = Format(rs!nMovImporte, gsFormatoNumeroView)
               fgDetalle.TextMatrix(n, 9) = rs!cCtaContCod
               If Not IsNull(rs!dItemPlazo) Then
                  fgDetalle.TextMatrix(n, 12) = rs!dItemPlazo
               End If
               If Not IsNull(rs!cDescrip) Then
                  fgDetalle.TextMatrix(n, 3) = rs!cDescrip
               End If
               'ALPA 20091110****************************************
               'fgDetalle.TextMatrix(n, 1) = Right(rs!cCtaContCod, 2)
               fgDetalle.TextMatrix(n, 1) = IIf(IsNull(rs!cAgeCod), Right(rs!cCtaContCod, 2), rs!cAgeCod) 'EJVG20140319
               '*****************************************************
            End If
            If Not IsNull(rs!cObjetoCod) Then
               Select Case rs!cObjetoCod
                  Case ObjCMACAgenciaArea, ObjCMACAgencias, ObjCMACArea
                     sSql = "SELECT nMovObjOrden, mo.cAreaCod+mo.cAgeCod cObjetoCod, " _
                          & "       ISNULL(ag.cAgeDescripcion,cAreaDescripcion) cObjetoDesc, NULL nMovCant " _
                          & "FROM MovObjAreaAgencia mo JOIN AreaAgencia aa ON aa.cAreaCod = mo.cAreaCod and aa.cAgeCod = mo.cAgeCod " _
                          & "     LEFT JOIN Areas a ON a.cAreaCod = aa.cAreaCod " _
                          & "     LEFT JOIN Agencias ag ON ag.cAgeCod = aa.cAgeCod " _
                          & "WHERE  mo.nMovNro = " & gcMovNro & " and mo.nMovItem = " & rs!nMovItem
                     sObjCod = Format(rs!cObjetoCod, "00")
                  Case ObjBienesServicios
                     'sSQL = "SELECT bs.nMovBsOrden nMovObjOrden, bs.cBSCod cObjetoCod, b.cBSDescripcion cObjetoDesc, mc.nMovCant " _
                     '     & "FROM " & sBaseNew & "MovBS bs JOIN " & sBaseNew & "MovCant mc ON mc.nMovNro = bs.nMovNro and mc.nMovItem = bs.nMovItem " _
                     '     & "     JOIN " & sBaseNew & "BienesServicios b ON b.cBSCod = bs.cBsCod " _
                     '     & "WHERE  bs.nMovNro = " & gcMovNro & " and bs.nMovItem = " & rs!nMovItem
                     'modificado por que no sale la unidad
                          
                      sSql = "SELECT bs.nMovBsOrden nMovObjOrden, bs.cBSCod cObjetoCod, " _
                            & " b.cBSDescripcion cObjetoDesc, mc.nMovCant ,rtrim(c.cConsDescripcion) as unidad " _
                            & " FROM MovBS bs ,MovCant mc ,BienesServicios b, constante c " _
                            & " Where mc.nMovNro = bs.nMovNro And mc.nMovItem = bs.nMovItem " _
                            & " and  b.cBSCod = bs.cBsCod and  b.nBSUnidad = c.nConsValor " _
                            & " and  bs.nMovNro = " & gcMovNro & " and bs.nMovItem = " & rs!nMovItem & " and  c.nConsCod = 1019 "
                          
                     sObjCod = Format(ObjBienesServicios, "00")
                  Case Else
                     sSql = "SELECT nMovObjOrden, mo.cObjetoCod, o.cObjetoDesc, NULL nMovCant " _
                          & "FROM " & sBaseNew & "MovObj mo JOIN  " & sBaseNew & "Objeto o ON o.cObjetoCod = mo.cObjetoCod " _
                          & "WHERE  mo.nMovNro = " & gcMovNro & " and mo.nMovItem = " & rs!nMovItem & " and not mo.cObjetoCod IN ('13','11','12','18') "
                     sObjCod = Mid(rs!cObjetoCod, 1, Len(rs!cObjetoCod))
               End Select
               Set rsObj = oCon.CargaRecordSet(sSql)
               If Not rsObj.EOF Then
                  If fgDetalle.TextMatrix(n, 3) = "" Then
                     fgDetalle.TextMatrix(n, 3) = rsObj!cObjetoDesc
                  End If
                  If Not IsNull(rsObj!nMovCant) Then
                     fgDetalle.TextMatrix(n, 2) = rsObj!cObjetoCod
                     fgDetalle.TextMatrix(n, 5) = rsObj!nMovCant
                     fgDetalle.TextMatrix(n, 4) = rsObj!unidad
                     If rs!nMovCant <> 0 Then
                        fgDetalle.TextMatrix(n, 6) = Format(Round(rs!nMovImporte / rsObj!nMovCant, 2), gcFormView)
                     End If
                     fgDetalle.TextMatrix(n, 11) = "B"
                     
                     FlexBackColor fgDetalle, n, lnColorBien
                     rs.MoveNext
                  Else
                     nItem = rs!nMovItem
                     fgDetalle.TextMatrix(n, 2) = rs!cCtaContCod
                     fgDetalle.TextMatrix(n, 11) = "S"
                     FlexBackColor fgDetalle, n, lnColorServ
                     'Comentado xPASIERS0472015 *****
'                     Do While Not rsObj.EOF
'                        If Not IsNull(rsObj!cObjetoCod) Then
'                           AdicionaRow fgObj
'                           fgObj.TextMatrix(fgObj.row, 0) = n
'                           fgObj.TextMatrix(fgObj.row, 2) = rs!cMovObjOrden
'                           '*** PEAC SE MODIFICÓ EL ORDEN
'                           fgObj.TextMatrix(fgObj.row, 3) = rsObj!cObjetoDesc 'rsObj!cObjetoCod
'                           fgObj.TextMatrix(fgObj.row, 4) = rsObj!cObjetoCod 'rsObj!cObjetoDesc
'                           fgObj.TextMatrix(fgObj.row, 5) = sObjCod
'                        End If
'                        rsObj.MoveNext
'                     Loop
                     'end PASI *****
                     rs.MoveNext
                  End If
               Else
                  rs.MoveNext
               End If
            Else
               fgDetalle.TextMatrix(n, 11) = "S"
               FlexBackColor fgDetalle, n, lnColorServ
               rs.MoveNext
            End If
         Loop
       
       SumasDoc
    Else
       Set rs = CargaOpeDocEstado(gcOpeCod, "1", "_2")
       If RSVacio(rs) Then
          MsgBox "No se definió Documento ORDEN DE COMPRA en Operación. Por favor Consultar con Sistemas...!", vbInformation, "¡Aviso!"
          lSalir = True
          Exit Sub
       Else
          gcDocTpo = rs!nDocTpo
          If Me.cmbmoneda.Text <> "" Then gcDocNro = ofun.GeneraDocNro(CInt(gcDocTpo), Right(Me.cmbmoneda.Text, 1), Year(gdFecSis), "OSC") 'NAGL 20191212 Agregó "OSC"
       End If
'       For N = 1 To fgDetalle.Rows - 1
'          If lbBienes Then
'             fgDetalle.TextMatrix(N, 11) = "B"
'          Else
'             fgDetalle.TextMatrix(N, 11) = "S"
'          End If
'       Next'EJVG20111115
       fgDetalle.row = 1
       fgDetalle.col = 2
       fgDetalle.row = 1: fgDetalle.col = 2
       txtPlazo.value = gdFecSis
    End If
    
    'If Me.cmbMoneda.Text <> "" Then gcDocNro = GeneraDocNro(CInt(gcDocTpo), Right(Me.cmbMoneda.Text, 1), Year(gdFecSis))
    
    LblDoc.Caption = UCase(sDocDesc) & "   Nº " & gcDocNro
    
    If lbBienes = True Then
       sTipodoc = "130"
       Else
       sTipodoc = "132"
    End If
    
    
    If gcMovNro <> "" Then
        sSql = " select  cDocNro from movdoc where nMovNro = " & gcMovNro & " and nDocTpo ='" & sTipodoc & "'"
        Set rs = oCon.CargaRecordSet(sSql)
        If rs.EOF Then
        Else
        lblDocOCD.Caption = UCase(sDocDesc) & " DIRECTA  Nº " & rs!cDocNro
        End If
    End If
      
    
    
    RSClose rs
    If lbImprime Or lbRegComp Then 'WIOR 20130110 AGREGO lbRegComp
       cmdDocumento.Visible = True
       cmdDocumentoPDF.Visible = True 'PASIERS1372014
       CmdAceptar.Visible = False
       cmdServicio.Visible = False
       chkProy.Enabled = False
       fraProy.Enabled = False
       fraProv.Enabled = False
       fraArea.Enabled = False
       fraLug.Enabled = False
       fraCambio.Enabled = False
       fraForm.Enabled = False
       fraMovDesc.Enabled = False
       fraCot.Enabled = False
       chkIGV.Visible = True
       If Not lbBienes Then
          chkRenta4.Visible = True
       End If
    End If
    
    Me.txtArea.rs = oArea.GetAgenciasAreas
    Me.txtCodLugarEntrega.rs = oConst.GetAgencias(, , True)
    txtPersona.TipoBusPers = BusPersDocumentoRuc
    fraCot.Visible = False
    'WIOR 20130110 *******************************
    If lbRegComp Then
        CmdAceptar.Visible = False
        chkIGV.Visible = False
        chkRenta4.Visible = False
        cmdServicio.Visible = False
        cmdDocumento.Visible = False
        cmdDocumentoPDF.Visible = False 'PASIERS1372014
        fraComprobante.Visible = True
        fraComprobante.Left = 7530
        If Mid(gcOpeCod, 3, 1) = 2 Then
            fraComprobante.Top = 1560
            fraForm.Width = 7275
            fraCambio.Height = 765
        Else
            fraComprobante.Top = 950
            fraForm.Width = 10875
            fraCambio.Height = 885
        End If
        
        cboDoc.Enabled = True
        txtFacNro.Enabled = True
        txtFacSerie.Enabled = True
        txtEmision.Enabled = True
        cmdDefinir.Enabled = True
        
        'Tipos de Comprobantes de Pago
        Set oDoc = New DOperacion
        Set rsDoc = oDoc.CargaOpeDoc(gcOpeCod, OpeDocMetDigitado)
        Do While Not rsDoc.EOF
           cboDoc.AddItem Format(rsDoc!nDocTpo, "00") & " - " & Mid(rsDoc!cDocDesc & Space(100), 1, 100) & rsDoc!nDocTpo
           rsDoc.MoveNext
        Loop
        Call CambiaTamañoCombo(cboDoc, 200)
        If cboDoc.ListCount = 1 Then
           cboDoc.ListIndex = 0
        End If
        RSClose rsDoc
        txtEmision.value = gdFecSis
    Else
        Set rsDoc = Nothing
'        cmdAceptar.Visible = True
'        chkIGV.Visible = True
'        chkRenta4.Visible = True
'        cmdDocumento.Visible = True
'        cmdServicio.Visible = True
        fraComprobante.Visible = False
        fraComprobante.Left = 11280
        fraComprobante.Top = 1080
        fraForm.Width = 10875
        fraCambio.Height = 885
    End If
    'WIOR FIN ************************************
End Sub

Private Sub textObjDes_GotFocus()
cmdExaminarDes.Visible = True
cmdExaminarDes.Top = textObjDes.Top + 15
cmdExaminarDes.Left = textObjDes.Left + txtObj.Width - cmdExaminarDes.Width
End Sub

Private Sub TxtArea_EmiteDatos()
    Me.txtAgeDesc.Text = txtArea.psDescripcion
    If txtArea.psDescripcion <> "" Then
        Me.txtCodLugarEntrega.SetFocus
    End If
End Sub

Private Sub TxtCodLugarEntrega_EmiteDatos()
    Dim oAge As DActualizaDatosArea
    Set oAge = New DActualizaDatosArea
    
    If txtCodLugarEntrega.Text <> "" Then
        Me.txtLugarEntrega.Text = oAge.GetDirAreaAgencia(, txtCodLugarEntrega)
        Me.txtLugarEntrega.SetFocus
    Else
        txtCodLugarEntrega.Text = ""
    End If
    fgDetalle.TextMatrix(fgDetalle.row, 1) = txtCodLugarEntrega.Text
End Sub

Private Sub txtConcepto_GotFocus()
fEnfoque txtConcepto
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   fgDetalle.TextMatrix(txtConcepto.Tag, 3) = txtConcepto.Text
   fgDetalle.SetFocus
End If
End Sub

Private Sub txtConcepto_LostFocus()
txtConcepto.Visible = False
End Sub

Private Sub txtConcepto_Validate(Cancel As Boolean)
fgDetalle.Text = txtConcepto.Text
End Sub

Private Sub txtCotNro_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   txtCodLugarEntrega.SetFocus
End If
End Sub



Private Sub txtDiasAtraso_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   KeyAscii = 0
   fgDetalle.SetFocus
End If
End Sub
'WIOR 20130110 *****************************************
Private Sub txtFacNro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       txtFacNro = Format(txtFacNro, String(8, "0"))
       txtEmision.SetFocus
    End If
End Sub

Private Sub txtFacNro_LostFocus()
  txtFacNro = Format(txtFacNro, String(8, "0"))
End Sub

Private Sub txtFacSerie_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
       txtFacSerie = Format(txtFacSerie, String(3, "0"))
       txtFacNro.SetFocus
    End If
End Sub

Private Sub txtFacSerie_LostFocus()
txtFacSerie = Format(txtFacSerie, String(3, "0"))
End Sub
'WIOR FIN  *********************************************
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFecha.Text) <> "" Then
      MsgBox " Fecha no Válida... ", vbInformation, "Aviso"
      txtFecha.SelStart = 0
      txtFecha.SelLength = Len(txtFecha)
   Else
      txtPersona.SetFocus
   End If
End If
End Sub

Private Sub txtFecha_Validate(Cancel As Boolean)
If ValidaFecha(txtFecha.Text) <> "" Then
   MsgBox " Fecha no Válida ", vbInformation, "Aviso"
   Cancel = True
End If
End Sub



Private Sub txtFecPlazo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFecPlazo) <> "" Then
      MsgBox "¡Fecha no Válida!", vbInformation, "¡Aviso!"
      txtFecPlazo.SetFocus
   Else
      fgDetalle.TextMatrix(txtFecPlazo.Tag, 12) = txtFecPlazo
      fgDetalle.SetFocus
   End If
End If
End Sub

Private Sub txtFecPlazo_LostFocus()
txtFecPlazo.Visible = False
End Sub

Private Sub txtLugarEntrega_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   txtMovDesc.SetFocus
End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   KeyAscii = 0
   txtDiasAtraso.SetFocus
End If
End Sub

Private Sub txtPersona_EmiteDatos()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oProv As DLogProveedor
    Set oProv = New DLogProveedor
    Me.txtProvNom.Text = txtPersona.psDescripcion
    lsCodPers = txtPersona.psCodigoPersona
    If txtPersona.psDescripcion <> "" Then
        Set rs = oProv.GetProveedorAgeRetBuenCont(lsCodPers)
        If rs.EOF And rs.BOF Then
            'MsgBox "La persona ingresada no esta registrada como proveedor o tiene el estado de Desactivado, debe regsitrarlo o activarlo.", vbInformation, "Aviso"
        Else
            Me.chkBuneCOnt.value = IIf(rs.Fields(1), 1, 0)
            Me.chkRetencion.value = IIf(rs.Fields(0), 1, 0)
        End If
        Me.txtArea.SetFocus
    End If
End Sub

Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   txtMovDesc.SetFocus
End If
End Sub

Private Sub txtPlazo_LostFocus()
Dim k As Integer
For k = 1 To fgDetalle.Rows - 1
   If fgDetalle.TextMatrix(k, 12) = "" And fgDetalle.TextMatrix(k, 2) <> "" Then
      fgDetalle.TextMatrix(k, 12) = txtPlazo
   Else
      If fgDetalle.TextMatrix(k, 2) = "" Then
         fgDetalle.TextMatrix(k, 12) = ""
      End If
   End If
Next
End Sub

Private Sub txtObj_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
   txtObj_KeyPress 13
   SendKeys IIf(KeyCode = 38, "{Up}", "{Down}"), True
End If
End Sub

Private Sub txtObj_GotFocus()
cmdExaminar.Visible = True
cmdExaminar.Top = txtObj.Top + 15
cmdExaminar.Left = txtObj.Left + txtObj.Width - cmdExaminar.Width
End Sub

Private Sub txtObj_KeyPress(KeyAscii As Integer)
Dim lbOK As Boolean
If KeyAscii = 13 Then
   If fgDetalle.TextMatrix(fgDetalle.row, 11) = "B" Then
      If ValidaObj Then
         ActualizaFG fgDetalle.row
         lbOK = True
      End If
   Else
      If ValidaCta Then
         fgDetalle.col = 8
         lbOK = True
      End If
   End If
   If lbOK Then
      fgDetalle.Enabled = True
      txtObj.Visible = False
      cmdExaminar.Visible = False
      fgDetalle.SetFocus
   End If
End If
End Sub
Private Function ValidaObj() As Boolean
Dim oCon As DConecta
Set oCon = New DConecta
oCon.AbreConexion

ValidaObj = False
If Len(txtObj) = 0 Then
   txtObj.Visible = False
   cmdExaminar.Visible = False
   Exit Function
End If

sSql = "Select cSubCtaCod from areaagencia where cAreaCod + cAgeCod = '" & Me.txtArea.Text & "'"
Set rs = oCon.CargaRecordSet(sSql)
sSql = FormaSelect(txtObj, 0, 0, rs.Fields(0))
Set rs = oCon.CargaRecordSet(sSql)
If Not RSVacio(rs) Then
   sObjCod = rs!cObjCod
   sObjDesc = rs!cObjDesc
   
   sObjUnid = IIf(IsNull(rs!cConsDescripcion) Or Trim(rs!cConsDescripcion) = "", "UND", rs!cConsDescripcion)
   If rs.RecordCount = 1 Then
      sCtaCod = rs!cObjetoCod   'Cuenta Contable
      sCtaDesc = rs!cObjetoDesc 'Cuenta Descrip
   Else
      frmDescObjeto.Inicio rs, "", 1, "Cuentas Contables"
      If frmDescObjeto.lOk Then
         sCtaCod = gaObj(0, 0, UBound(gaObj, 3))
         sCtaDesc = gaObj(0, 1, UBound(gaObj, 3))
      Else
         MsgBox "No se definió Cuenta Contable", vbInformation, "Aviso"
         Exit Function
      End If
   End If
Else
   MsgBox "Objeto no encontrado...!", vbInformation, "!Aviso¡"
   Exit Function
End If
ValidaObj = True
End Function

Private Sub txtCant_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCant, KeyAscii, 12, 2)
If KeyAscii = 13 Then
   fgDetalle.Text = Format(txtCant.Text, gcFormView)
   If fgDetalle.col = 5 Then
      fgDetalle.TextMatrix(fgDetalle.row, 8) = Format(Round(Val(txtCant) * Val(Format(fgDetalle.TextMatrix(fgDetalle.row, 6), "#0.00")), 2), gcFormView)
   End If
   If fgDetalle.col = 6 Then
      fgDetalle.TextMatrix(fgDetalle.row, 8) = Format(Round(Val(txtCant) * Val(Format(fgDetalle.TextMatrix(fgDetalle.row, 5), "#0.00")), 2), gcFormView)
   End If
   If fgDetalle.col = 8 And Val(fgDetalle.TextMatrix(fgDetalle.row, 5)) <> 0 Then
      fgDetalle.TextMatrix(fgDetalle.row, 6) = Format(Round(Val(txtCant) / Val(Format(fgDetalle.TextMatrix(fgDetalle.row, 5), "#0.00")), 2), gcFormView)
   End If
   txtCant.Visible = False
   fgDetalle.SetFocus
   SumasDoc
End If
End Sub
Private Sub SumasDoc()
Dim n As Integer
Dim nTot As Currency
For n = 1 To fgDetalle.Rows - 1
    If fgDetalle.TextMatrix(n, 8) <> "" Then
       nTot = nTot + Val(Format(fgDetalle.TextMatrix(n, 8), gcFormDato))
    End If
Next
If nTot > 0 Then
   txtTot = Format(nTot, gcFormView)
Else
   txtTot = ""
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
If Val(txtCant) = 0 Then
   MsgBox "Cantidad debe ser mayor que Cero...!", vbCritical, "Aviso"
   Cancel = True
End If
End Sub

Private Sub txtTipCompra_GotFocus()
    txtTipCompra.SelStart = 0
    txtTipCompra.SelLength = 50
End Sub

Private Sub txtTipCompra_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCompra, KeyAscii, 12, 4)
If KeyAscii = 13 Then
   txtMovDesc.SetFocus
End If
End Sub

'Private Sub cmdAceptar_Click()
'    Dim n As Integer, m As Integer  'Contador
'    Dim nItem As Integer, nCol  As Integer
'    Dim sTexto As String, lOk   As Boolean
'    Dim nObj As Integer
'    Dim lsMovNro As String
'    Dim lnMovNro As Long
'    Dim lnOrdenObj As Integer
'    Dim oMov As DMov
'    Dim lsAreaCod As String
'    Dim lsAgeCod As String
'    Dim lsDocNroOCD As String
'    Set oMov = New DMov
'    Dim sTipodoc As String
'    Dim sPrenombre As String
'    Dim ofun As NContFunciones
'    Set ofun = New NContFunciones
'
'    On Error GoTo ErrAceptar
'    'On Error Resume Next
'
'    If Not ValidaDatos() Then
'       Exit Sub
'    End If
'
'    If cmbTipoOC.Text = "" Then
'       MsgBox "Debe seleccionar si la orden  es Directa o con Proceso", vbInformation, "Aviso"
'       Exit Sub
'    End If
'
'    If MsgBox(" ¿ Esta usted Seguro de grabar la Operación ? ", vbQuestion + vbYesNo, "Confirmacion") = vbNo Then
'       Exit Sub
'    End If
'
'    If Right(cmbTipoOC.Text, 1) = gLogOCDirecta Then
'       sPrenombre = " DIRECTA"
'    ElseIf Right(cmbTipoOC.Text, 1) = gLogOCProceso Then
'       sPrenombre = " PROCESO"
'    End If
'
'
'
'    cmdAceptar.Enabled = False
'    ' Iniciamos transaccion
'    If lTransActiva Then
'       oMov.RollbackTrans
'       lTransActiva = False
'    End If
'    oMov.BeginTrans
'    lTransActiva = True
'
'    If Not lbModifica Then
'        gcDocNro = oMov.GeneraDocNro(CInt(gcDocTpo), Mid(gcOpeCod, 3, 1), Year(gdFecSis))
'        lblDoc.Caption = UCase(sDocDesc) & " Nº " & gcDocNro
'        'if pasa el tope entonces
'        If lbBienes = True Then
'           If Right(cmbTipoOC.Text, 1) = gLogOCDirecta Then
'                    sTipodoc = LogTipoOC.gLogOCompraDirecta
'               ElseIf Right(cmbTipoOC.Text, 1) = gLogOCProceso Then
'                    sTipodoc = LogTipoOC.gLogOCompraProceso
'           End If
'
'        Else
'            If Right(cmbTipoOC.Text, 1) = gLogOCDirecta Then
'                    sTipodoc = LogTipoOC.gLogOServicioDirecta
'            ElseIf Right(cmbTipoOC.Text, 1) = gLogOCProceso Then
'                    sTipodoc = LogTipoOC.gLogOServicioProceso
'            End If
'        End If
'
'
'
'        If Right(cmbTipoOC.Text, 1) = gLogOCDirecta Or Right(cmbTipoOC.Text, 1) = gLogOCProceso Then
'                lsDocNroOCD = ofun.GeneraDocNro(CInt(sTipodoc), Mid(gcOpeCod, 3, 1), Year(gdFecSis))
'                lblDocOCD.Caption = UCase(sDocDesc) & sPrenombre & " Nº " & lsDocNroOCD
'            If MsgBox(" ¿ Esta creando una " & sDocDesc & sPrenombre & "  desea Continuar ? ", vbQuestion + vbYesNo, "") = vbNo Then
'                If lTransActiva Then
'                   oMov.RollbackTrans
'                End If
'                lTransActiva = False
'                cmdAceptar.Enabled = True
'                Exit Sub
'            End If
'
'        End If
'
'
'    End If
'
'
'    lsMovNro = oMov.GeneraMovNro(gdFecSis, , gsCodUser)
'    If lbModifica Then
'        oMov.ActualizaEstadoMov gcMovNro, gMovFlagModificado
'        oMov.InsertaMovModifica oMov.GeneraMovNro(gdFecSis, , gsCodUser), oMov.GetcMovNro(gcMovNro), lsMovNro
'    End If
'
'    ' Grabamos en Mov
'    If sProvCod = "" Then
'        gcMovNro = oMov.GeneraMovNro(gdFecSis, , gsCodUser)
'    Else
'        gcMovNro = oMov.GeneraMovNro(gdFecSis, , gsCodUser, oMov.GetcMovNro(gcMovNro))
'    End If
'    gsGlosa = txtMovDesc
'
'    oMov.InsertaMov gcMovNro, gcOpeCod, gsGlosa, gMovEstPresupPendiente
'    lnMovNro = oMov.GetnMovNro(gcMovNro)
'    lsMovNro = gcMovNro
'    gcMovNro = lnMovNro
'
'    ' Grabamos en MovObj y MovCant
'    nItem = 0
'    gnImporte = 0
'    For n = 1 To fgDetalle.Rows - 1
'       If Len(fgDetalle.TextMatrix(n, 1)) > 0 And nVal(fgDetalle.TextMatrix(n, 7)) > 0 Then
'          nItem = nItem + 1
'          sTexto = ""
'          If fgDetalle.TextMatrix(n, 10) = "B" Then
'             oMov.InsertaMovCta lnMovNro, nItem, fgDetalle.TextMatrix(n, 8) & sTexto, nVal(fgDetalle.TextMatrix(n, 7))
'             oMov.InsertaMovObj lnMovNro, nItem, 1, ObjBienesServicios
'             oMov.InsertaMovBS lnMovNro, nItem, 1, fgDetalle.TextMatrix(n, 1)
'          Else
'             nObj = 1
'             'Recorrido para formar cuenta
'             For m = 1 To fgObj.Rows - 1
'                If fgObj.TextMatrix(m, 0) = fgDetalle.TextMatrix(n, 0) Then
'                    sTexto = sTexto & fgObj.TextMatrix(m, 4)
'                End If
'             Next m
'             oMov.InsertaMovCta lnMovNro, nItem, fgDetalle.TextMatrix(n, 8) & sTexto, nVal(fgDetalle.TextMatrix(n, 7))
'
'             'Graba Objetos
'             lnOrdenObj = 0
'             For m = 1 To fgObj.Rows - 1
'                If fgObj.TextMatrix(m, 0) = fgDetalle.TextMatrix(n, 0) And fgObj.TextMatrix(m, 1) <> "" Then
'
'                    Select Case fgObj.TextMatrix(m, 5)
'                        Case ObjCMACAgencias
'                            lnOrdenObj = lnOrdenObj + 1
'                            lsAgeCod = fgObj.TextMatrix(m, 2)
'                            oMov.InsertaMovObj lnMovNro, nItem, lnOrdenObj, ObjCMACAgencias
'                            oMov.InsertaMovObjAgenciaArea lnMovNro, nItem, lnOrdenObj, lsAgeCod, ""
'                        Case ObjCMACArea
'                            lnOrdenObj = lnOrdenObj + 1
'                            lsAreaCod = Mid(fgObj.TextMatrix(m, 2), 1, 3)
'                            oMov.InsertaMovObj lnMovNro, nItem, lnOrdenObj, ObjCMACArea
'                            oMov.InsertaMovObjAgenciaArea lnMovNro, nItem, lnOrdenObj, "", lsAreaCod
'                        Case ObjCMACAgenciaArea
'                            lsAreaCod = Mid(fgObj.TextMatrix(m, 2), 1, 3)
'                            lsAgeCod = Mid(fgObj.TextMatrix(m, 2), 4, 2)
'                            lnOrdenObj = lnOrdenObj + 1
'                            oMov.InsertaMovObj lnMovNro, nItem, lnOrdenObj, ObjCMACAgenciaArea
'                            oMov.InsertaMovObjAgenciaArea lnMovNro, nItem, lnOrdenObj, lsAgeCod, lsAreaCod
'                        Case ObjPersona
'                            lnOrdenObj = lnOrdenObj + 1
'                            oMov.InsertaMovObj lnMovNro, nItem, lnOrdenObj, ObjPersona
'                            oMov.InsertaMovGasto lnMovNro, lsCodPers, ""
'                        Case ObjBienesServicios
'                            lnOrdenObj = lnOrdenObj + 1
'                            oMov.InsertaMovObj lnMovNro, nItem, lnOrdenObj, ObjBienesServicios
'                            oMov.InsertaMovBS lnMovNro, nItem, lnOrdenObj, fgObj.TextMatrix(m, 2)
'                        Case Else
'                            lnOrdenObj = lnOrdenObj + 1
'                            oMov.InsertaMovObj lnMovNro, nItem, lnOrdenObj, fgObj.TextMatrix(m, 2)
'                    End Select
'                End If
'             Next
'          End If
'
'          'Grabamos el detalle del Bien o Servicio
'          oMov.InsertaMovCotizacDet lnMovNro, nItem, fgDetalle.TextMatrix(n, 2), Format(CDate(fgDetalle.TextMatrix(n, 11)), gsFormatoFecha)
'          gnImporte = gnImporte + Val(Format(fgDetalle.TextMatrix(n, 7), gcFormDato))
'
'          'Si cantidad > 0 then
'          If fgDetalle.TextMatrix(n, 4) <> "" Then
'                If nVal(fgDetalle.TextMatrix(n, 4)) <> 0 Then
'                     'Actualizamos los Stocks
'                      oMov.InsertaMovCant lnMovNro, nItem, nVal(fgDetalle.TextMatrix(n, 4))
'                End If
'           End If
'       End If
'    Next
'    'Grabamos la Provisión
'    If gnImporte = 0 Then
'       Err.Raise "50001", "Grabando", "No se puede grabar Documento sin Importe"
'    End If
'    nItem = nItem + 1
'    oMov.InsertaMovCta lnMovNro, nItem, sCtaProvis, gnImporte * -1
'    oMov.InsertaMovGasto lnMovNro, lsCodPers, "1"
'    ' Actualizamos Nro de Orden de Compra
'
'
'
'    If Right(cmbTipoOC.Text, 1) = gLogOCDirecta Or Right(cmbTipoOC.Text, 1) = gLogOCProceso Then
'
'       If lbModifica = True Then
'            lsDocNroOCD = Right(lblDocOCD.Caption, 13)
'       End If
'
'       If lbBienes = True Then
'           If Right(cmbTipoOC.Text, 1) = gLogOCDirecta Then
'                    sTipodoc = LogTipoOC.gLogOCompraDirecta
'               ElseIf Right(cmbTipoOC.Text, 1) = gLogOCProceso Then
'                    sTipodoc = LogTipoOC.gLogOCompraProceso
'           End If
'        Else
'            If Right(cmbTipoOC.Text, 1) = gLogOCDirecta Then
'                    sTipodoc = LogTipoOC.gLogOServicioDirecta
'            ElseIf Right(cmbTipoOC.Text, 1) = gLogOCProceso Then
'                    sTipodoc = LogTipoOC.gLogOServicioProceso
'            End If
'        End If
'
'       If lsDocNroOCD <> "" Then
'            oMov.InsertaMovDoc lnMovNro, sTipodoc, lsDocNroOCD, Format(txtFecha, gsFormatoFecha)
'       End If
'    End If
'
'    oMov.InsertaMovDoc lnMovNro, gcDocTpo, gcDocNro, Format(txtFecha, gsFormatoFecha)
'
'
'    oMov.InsertaMovCotizac lnMovNro, txtCotNro.Text, txtArea.Text, Right(cboFormaPago, 1), Right(Me.cmbMoneda.Text, 1), Format(txtPlazo, gsFormatoFecha), txtLugarEntrega.Text, txtProy.Tag, Right(cmbTipoOC.Text, 1), Val(Trim(txtDiasAtraso.Text))
'
'    'Grabamos el Tipo de Cambio de Compra
'    If Right(cmbMoneda.Text, 1) = gcMEDig Then
'        'oMov.InsertaMovMe lnMovNro, 1, Format(txtTipCompra, "########0.0000")
'        oMov.GeneraMovME lnMovNro, lsMovNro, Format(txtTipCompra, "########0.0000")
'    End If
'
'    oMov.CommitTrans
'    lTransActiva = False
'       OK = True
'       If lOrdenCompra And Not lbModifica And Not lbImprime Then
'          If MsgBox(" ¿ Desea Ingresar nuevo Documento ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbYes Then
'             gcDocNro = ofun.GeneraDocNro(Val(gcDocTpo), Right(Me.cmbMoneda.Text, 1), Year(gdFecSis))
'             For n = 1 To fgDetalle.Rows - 1
'                For nItem = 1 To fgDetalle.Cols - 1
'                   fgDetalle.TextMatrix(n, nItem) = ""
'                Next
'             Next
'             fgObj.Rows = 2
'             For nItem = 1 To fgObj.Cols - 1
'                fgObj.TextMatrix(1, nItem) = ""
'             Next
'             txtTot = ""
'             txtMovDesc = ""
'             lblDoc.Caption = UCase(sDocDesc) & "    Nº " & gcDocNro
'
'
'             lsDocNroOCD = ofun.GeneraDocNro(CInt(sTipodoc), Mid(gcOpeCod, 3, 1), Year(gdFecSis))
'             lblDocOCD.Caption = UCase(sDocDesc) & sPrenombre & " Nº " & lsDocNroOCD
'             fgDetalle.SetFocus
'          Else
'             Unload Me
'          End If
'       Else
'          Unload Me
'       End If
'
'    cmdAceptar.Enabled = True
'    Exit Sub
'ErrAceptar:
'    If lTransActiva Then
'        oMov.RollbackTrans
'    End If
'    lTransActiva = False
'    cmdAceptar.Enabled = True
'    MsgBox TextErr(Err.Description), vbCritical, "Error de Actualización"
'End Sub

Private Sub CmdAceptar_Click()
    Dim n As Integer, m As Integer  'Contador
    Dim nItem As Integer, nCol  As Integer
    Dim sTexto As String, lOk   As Boolean
    Dim nObj As Integer
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lnOrdenObj As Integer
    Dim oMov As DMov
    Dim lsAreaCod As String
    Dim lsAgeCod As String
    Dim nTipoCambioC As Currency
    Dim I As Integer
    Dim nn As Integer
    Set oMov = New DMov
    Dim sTipodoc As String
    Dim sPreNombre As String
    Dim lsAgeCodItem As String 'EJVG20140319
    Dim olog As DLogGeneral 'PASI20150918 ERS0472015
    Set olog = New DLogGeneral
    Dim nx As Integer
    Dim lsMontoContrato() As Currency
    Dim k As Integer
    Dim bestado As Boolean
    Dim nMovNroComp As Long
    Dim nMovNroOrdAnt As Long
    'END PASI
    
    Dim lsMovNroAux As String
    
    'Inicio - Modificado por ORCR 03/02/2014
    Dim lsCtaContCod As String
    'Fin - Modificado por ORCR 03/02/2014
    
    On Error GoTo ErrAceptar    'On Error Resume Next
    
           
    If Not ValidaDatos() Then
       Exit Sub
    End If
    
    If cmbTipoOC.Text = "" Then
       MsgBox "Debe seleccionar si la orden  es Directa o con Proceso", vbInformation, "Aviso"
       Exit Sub
    End If
   
    If MsgBox(" ¿ Seguro de grabar Operación ? ", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
       Exit Sub
    End If
    If Right(cmbTipoOC.Text, 1) = gLogOCDirecta Then
       sPreNombre = " DIRECTA"
    ElseIf Right(cmbTipoOC.Text, 1) = gLogOCProceso Then
       sPreNombre = " PROCESO"
    End If
   
    CmdAceptar.Enabled = False
    ' Iniciamos transaccion
    If lTransActiva Then
    oMov.RollbackTrans
    lTransActiva = False
    End If
    oMov.BeginTrans
    lTransActiva = True
        
        If Not lbModifica Then
            gcDocNro = oMov.GeneraDocNro(CInt(gcDocTpo), Mid(gcOpeCod, 3, 1), Year(gdFecSis), "OSC") 'NAGL 20191212 Agregó "OSC"
            LblDoc.Caption = UCase(sDocDesc) & " Nº " & gcDocNro
            
            If lbBienes = True Then
                If Right(cmbTipoOC.Text, 1) = gLogOCDirecta Then
                    sTipodoc = LogTipoOC.gLogOCompraDirecta
                ElseIf Right(cmbTipoOC.Text, 1) = gLogOCProceso Then
                    sTipodoc = LogTipoOC.gLogOCompraProceso
                End If
           
            Else
                If Right(cmbTipoOC.Text, 1) = gLogOCDirecta Then
                    sTipodoc = LogTipoOC.gLogOServicioDirecta
                ElseIf Right(cmbTipoOC.Text, 1) = gLogOCProceso Then
                    sTipodoc = LogTipoOC.gLogOServicioProceso
                End If
            End If
        
               
            If Right(cmbTipoOC.Text, 1) = gLogOCDirecta Or Right(cmbTipoOC.Text, 1) = gLogOCProceso Then
                    lsDocNroOCD = oMov.GeneraDocNro(CInt(sTipodoc), Mid(gcOpeCod, 3, 1), Year(gdFecSis), "OSC") 'NAGL 20191212 Agregó "OSC"
                    lblDocOCD.Caption = UCase(sDocDesc) & sPreNombre & " Nº " & lsDocNroOCD
                If MsgBox(" ¿ Esta creando una " & sDocDesc & sPreNombre & "  desea Continuar ? ", vbQuestion + vbYesNo, "") = vbNo Then
                    If lTransActiva Then
                       oMov.RollbackTrans
                    End If
                    lTransActiva = False
                    CmdAceptar.Enabled = True
                    Exit Sub
                End If
        
            End If
        End If
        
        lsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        If lbModifica Then
            'oMov.ActualizaEstadoMov gcMovNro, gMovFlagModificado
            'oMov.InsertaMovModifica oMov.GeneraMovNro(gdFecSis, , gsCodUser), oMov.GetcMovNro(gcMovNro), lsMovNro
            'PASI2051210********
            oMov.ActualizaEstadoMov lsMovNroEdit, gMovFlagModificado
            oMov.InsertaMovModifica oMov.GeneraMovNro(gdFecSis, , gsCodUser), oMov.GetcMovNro(lsMovNroEdit), lsMovNro
            lsMovNroAux = lsMovNroEdit
            '*******************
        End If
        If Not lsMovNroEdit = "" Then nMovNroOrdAnt = lsMovNroEdit 'PASI20150918 ERS0472015
        ' Grabamos en Mov
        If sProvCod = "" Then
            gcMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Else
            'gcMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser, oMov.GetcMovNro(gcMovNro))
            gcMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser, oMov.GetcMovNro(lsMovNroEdit)) 'PASI20151210
        End If
        
        'lsMovNroAux = gcMovNro
        gsGlosa = txtMovDesc
        'lsMovNroAux = ""
        
        oMov.InsertaMov gcMovNro, gcOpeCod, gsGlosa, gMovEstPresupAceptado 'PASI20150918 ERS0472015 cambio el parametro  gMovEstPresupPendiente por gMovEstPresupAceptado
                
        lnMovNro = oMov.GetnMovNro(gcMovNro)
        lsMovNro = gcMovNro
        gcMovNro = lnMovNro
        
        'ARLO 20170125
        Dim sOpe As String
        If lbModifica Then
        sOpe = 2
        Else
        sOpe = 1
        End If
        If (cmbmoneda.ListIndex = 1) And sDocDesc = "Orden de Servicio" Then
        gsOpeCod = LogPistaOrdenServicioSoles
        ElseIf (cmbmoneda.ListIndex = 1) And sDocDesc = "Orden de Compra" Then
        gsOpeCod = LogPistaOrdenCompraSoles
        ElseIf (cmbmoneda.ListIndex <> 1) And sDocDesc = "Orden de Servicio" Then
        gsOpeCod = LogPistaOrdenServicioSoles
        Else: gsOpeCod = LogPistaOrdenCompraDolares
        End If
        
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, sOpe, sDocDesc & " " & sPreNombre & " en " & Mid(cmbmoneda.Text, 1, 7) & " N°: " & gcDocNro
        Set objPista = Nothing
        '***********
        
        ' Grabamos en MovObj y MovCant
        nItem = 0
        gnImporte = 0
        For n = 1 To fgDetalle.Rows - 1
           If Len(fgDetalle.TextMatrix(n, 2)) > 0 And nVal(fgDetalle.TextMatrix(n, 8)) > 0 Then
              nItem = nItem + 1
              sTexto = ""
              lsAgeCodItem = Format(fgDetalle.TextMatrix(n, 1), "00") 'EJVG20140319
              oMov.InsertaMovAgencia lnMovNro, nItem, lsAgeCodItem 'EJVG20140319
              If fgDetalle.TextMatrix(n, 11) = "B" Then
                 oMov.InsertaMovCta lnMovNro, nItem, fgDetalle.TextMatrix(n, 9) & sTexto, nVal(fgDetalle.TextMatrix(n, 8))
                 oMov.InsertaMovObj lnMovNro, nItem, 1, ObjBienesServicios
                 oMov.InsertaMovBS lnMovNro, nItem, CLng(lsAgeCodItem), fgDetalle.TextMatrix(n, 2) 'EJVG20140319
              Else
                 nObj = 1
                 'Recorrido para formar cuenta
                 For m = 1 To fgObj.Rows - 1
                    If fgObj.TextMatrix(m, 0) = fgDetalle.TextMatrix(n, 0) Then
                        sTexto = sTexto & fgObj.TextMatrix(m, 4)
                    End If
                 Next m
                'EJVG20140225 ***
                
                'Modificado PASI20150917 ERS0472015
'                If lbModifica = True Then
'                    'lsCtaContCod = Left(fgDetalle.TextMatrix(n, 9), Len(fgDetalle.TextMatrix(n, 9)) - 2) & fgDetalle.TextMatrix(n, 1)
'                    lsCtaContCod = Left(fgDetalle.TextMatrix(n, 9), Len(fgDetalle.TextMatrix(n, 9)) - 2) & lsAgeCodItem 'EJVG20140319
'                    oMov.InsertaMovCta lnMovNro, nItem, lsCtaContCod, nVal(fgDetalle.TextMatrix(n, 8))
'                Else
'                    oMov.InsertaMovCta lnMovNro, nItem, fgDetalle.TextMatrix(n, 9) & sTexto, nVal(fgDetalle.TextMatrix(n, 8))
'                End If
                
'                If lbModifica Then
'                    oMov.InsertaMovCta lnMovNro, nItem, fgDetalle.TextMatrix(n, 9), nVal(fgDetalle.TextMatrix(n, 8))
'                Else
                    oMov.InsertaMovCta lnMovNro, nItem, fgDetalle.TextMatrix(n, 9) & sTexto, nVal(fgDetalle.TextMatrix(n, 8))
'                End If
                
                'END PASI
           
                 'Graba Objetos
                 lnOrdenObj = 0
                 For m = 1 To fgObj.Rows - 1
                    If fgObj.TextMatrix(m, 0) = fgDetalle.TextMatrix(n, 0) And fgObj.TextMatrix(m, 1) <> "" Then
                     
                        Select Case fgObj.TextMatrix(m, 5)
                            Case ObjCMACAgencias
                                lnOrdenObj = lnOrdenObj + 1
                                lsAgeCod = fgObj.TextMatrix(m, 2)
                                oMov.InsertaMovObj lnMovNro, nItem, lnOrdenObj, ObjCMACAgencias
                                oMov.InsertaMovObjAgenciaArea lnMovNro, nItem, lnOrdenObj, lsAgeCod, ""
                            Case ObjCMACArea
                                lnOrdenObj = lnOrdenObj + 1
                                lsAreaCod = Mid(fgObj.TextMatrix(m, 2), 1, 3)
                                oMov.InsertaMovObj lnMovNro, nItem, lnOrdenObj, ObjCMACArea
                                oMov.InsertaMovObjAgenciaArea lnMovNro, nItem, lnOrdenObj, "", lsAreaCod
                            Case ObjCMACAgenciaArea
                                lsAreaCod = Mid(fgObj.TextMatrix(m, 2), 1, 3)
                                lsAgeCod = Mid(fgObj.TextMatrix(m, 2), 4, 2)
                                lnOrdenObj = lnOrdenObj + 1
                                oMov.InsertaMovObj lnMovNro, nItem, lnOrdenObj, ObjCMACAgenciaArea
                                oMov.InsertaMovObjAgenciaArea lnMovNro, nItem, lnOrdenObj, lsAgeCod, lsAreaCod
                            Case ObjPersona
                                lnOrdenObj = lnOrdenObj + 1
                                oMov.InsertaMovObj lnMovNro, nItem, lnOrdenObj, ObjPersona
                                oMov.InsertaMovGasto lnMovNro, lsCodPers, ""
                            Case ObjBienesServicios
                                lnOrdenObj = lnOrdenObj + 1
                                oMov.InsertaMovObj lnMovNro, nItem, lnOrdenObj, ObjBienesServicios
                                oMov.InsertaMovBS lnMovNro, nItem, lnOrdenObj, fgObj.TextMatrix(m, 2)
                            Case Else
                                lnOrdenObj = lnOrdenObj + 1
                                oMov.InsertaMovObj lnMovNro, nItem, lnOrdenObj, Val(fgObj.TextMatrix(m, 2))
                        End Select
'                    Else
'                        lnOrdenObj = lnOrdenObj + 1
'                        oMov.InsertaMovObj lnMovNro, nItem, lnOrdenObj, ObjBienesServicios
                    End If
                 Next
                 
                 'PASIERS0472015 para Regularizar OBJs *************
                    If lbModifica Then
                        oMov.RegularizaMovObj lsMovNroAux, lnMovNro, nItem, fgDetalle.TextMatrix(n, 9) & sTexto
                    End If
                 'end PASI *******
                 
                 
              End If
              
              'Grabamos el detalle del Bien o Servicio
              oMov.InsertaMovCotizacDet lnMovNro, nItem, fgDetalle.TextMatrix(n, 3), Format(CDate(fgDetalle.TextMatrix(n, 12)), gsFormatoFecha)
              gnImporte = gnImporte + Val(Format(fgDetalle.TextMatrix(n, 8), gcFormDato))
              
              'Si cantidad > 0 then
              If fgDetalle.TextMatrix(n, 5) <> "" Then
                    If nVal(fgDetalle.TextMatrix(n, 5)) <> 0 Then
                         'Actualizamos los Stocks
                          oMov.InsertaMovCant lnMovNro, nItem, nVal(fgDetalle.TextMatrix(n, 5))
                    End If
               End If
           End If
        Next
        'Grabamos la Provisión
        If gnImporte = 0 Then
           Err.Raise "50001", "Grabando", "No se puede grabar Documento sin Importe"
        End If
        nItem = nItem + 1
        oMov.InsertaMovCta lnMovNro, nItem, sCtaProvis, gnImporte * -1
        oMov.InsertaMovGasto lnMovNro, lsCodPers, "1"
        
        If Right(cmbTipoOC.Text, 1) = gLogOCDirecta Or Right(cmbTipoOC.Text, 1) = gLogOCProceso Then
       
       If lbModifica = True Then
            lsDocNroOCD = Right(lblDocOCD.Caption, 13)
       End If
       
       If lbBienes = True Then
           If Right(cmbTipoOC.Text, 1) = gLogOCDirecta Then
                    sTipodoc = LogTipoOC.gLogOCompraDirecta
               ElseIf Right(cmbTipoOC.Text, 1) = gLogOCProceso Then
                    sTipodoc = LogTipoOC.gLogOCompraProceso
           End If
        Else
            If Right(cmbTipoOC.Text, 1) = gLogOCDirecta Then
                    sTipodoc = LogTipoOC.gLogOServicioDirecta
            ElseIf Right(cmbTipoOC.Text, 1) = gLogOCProceso Then
                    sTipodoc = LogTipoOC.gLogOServicioProceso
            End If
        End If
        
       If lsDocNroOCD <> "" Then
            oMov.InsertaMovDoc lnMovNro, sTipodoc, lsDocNroOCD, Format(txtFecha, gsFormatoFecha)
       End If
    End If
    
        ' Actualizamos Nro de Orden de Compra
          
        oMov.InsertaMovDoc lnMovNro, gcDocTpo, gcDocNro, Format(txtFecha, gsFormatoFecha)
        oMov.InsertaMovCotizac lnMovNro, txtCotNro.Text, txtArea.Text, Right(cboFormaPago, 1), Right(Me.cmbmoneda.Text, 1), Format(txtPlazo, gsFormatoFecha), txtLugarEntrega.Text, txtProy.Tag, Right(cmbTipoOC.Text, 1), Val(Trim(txtDiasAtraso.Text))
        'Grabamos el Tipo de Cambio de Compra
        If Right(cmbmoneda.Text, 1) = gcMEDig Then
            'oMov.InsertaMovMe lnMovNro, 1, Format(txtTipCompra, "########0.0000")
            oMov.GeneraMovME lnMovNro, lsMovNro, Format(txtTipCompra, "########0.0000"), True
        End If
        If chkContSuministro.value = 1 Then 'PASI20140818 ERS0772014
            If sProvCod = "" Then
                    ReDim Preserve lsMontoContrato(3, 0)
                    For nx = 1 To fgDetalle.Rows - 1
                        If fgDetalle.TextMatrix(nx, 1) <> "" And fgDetalle.TextMatrix(nx, 2) <> "" Then
                            If nx = 1 Then
                                ReDim Preserve lsMontoContrato(3, 1)
                                lsMontoContrato(1, 1) = fgDetalle.TextMatrix(nx, 2)
                                lsMontoContrato(2, 1) = fgDetalle.TextMatrix(nx, 5)
                                lsMontoContrato(3, 1) = fgDetalle.TextMatrix(nx, 8)
                            Else
                                For k = 1 To UBound(lsMontoContrato, 2)
                                    If fgDetalle.TextMatrix(nx, 2) = lsMontoContrato(1, k) Then
                                        lsMontoContrato(2, k) = lsMontoContrato(2, k) + fgDetalle.TextMatrix(nx, 5)
                                        lsMontoContrato(3, k) = lsMontoContrato(3, k) + fgDetalle.TextMatrix(nx, 8)
                                        bestado = True
                                    End If
                                Next
                                If Not bestado Then
                                    ReDim Preserve lsMontoContrato(3, UBound(lsMontoContrato, 2) + 1)
                                    lsMontoContrato(1, UBound(lsMontoContrato, 2)) = fgDetalle.TextMatrix(nx, 2)
                                    lsMontoContrato(2, UBound(lsMontoContrato, 2)) = fgDetalle.TextMatrix(nx, 5)
                                    lsMontoContrato(3, UBound(lsMontoContrato, 2)) = fgDetalle.TextMatrix(nx, 8)
                                End If
                                bestado = False
                            End If
                        End If
                    Next
                    For k = 1 To UBound(lsMontoContrato, 2)
                        oMov.InsertaContratoBienOrden Trim(Right(cboContSuministro.Text, 100)), lsMontoContrato(1, k), lsMontoContrato(2, k), lsMontoContrato(3, k), lsMovNro
                    Next
            Else
                    ReDim Preserve lsMontoContrato(3, 0)
                    For nx = 1 To fgDetalle.Rows - 1
                        If n = 1 Then
                            ReDim Preserve lsMontoContrato(3, 1)
                            lsMontoContrato(1, 1) = fgDetalle.TextMatrix(nx, 2)
                            lsMontoContrato(2, 1) = fgDetalle.TextMatrix(nx, 5)
                            lsMontoContrato(3, 1) = fgDetalle.TextMatrix(nx, 8)
                        Else
                            For k = 1 To UBound(lsMontoContrato, 2)
                                If fgDetalle.TextMatrix(nx, 2) = lsMontoContrato(1, k) Then
                                    lsMontoContrato(2, 1) = lsMontoContrato(2, k) + fgDetalle.TextMatrix(nx, 5)
                                    lsMontoContrato(3, k) = lsMontoContrato(2, k) + fgDetalle.TextMatrix(n, 8)
                                    bestado = True
                                End If
                            Next
                            If Not bestado Then
                                ReDim Preserve lsMontoContrato(3, UBound(lsMontoContrato, 2) + 1)
                                lsMontoContrato(1, UBound(lsMontoContrato, 2)) = fgDetalle.TextMatrix(nx, 2)
                                lsMontoContrato(2, UBound(lsMontoContrato, 2)) = fgDetalle.TextMatrix(nx, 5)
                                lsMontoContrato(3, UBound(lsMontoContrato, 2)) = fgDetalle.TextMatrix(nx, 8)
                            End If
                            bestado = False
                        End If
                    Next
                    For k = 1 To UBound(lsMontoContrato, 2)
                        oMov.InsertaContratoBienOrdenEdit Trim(Right(cboContSuministro.Text, 100)), lsMontoContrato(1, k), lsMontoContrato(2, k), lsMontoContrato(3, k), lsMovNro, oMov.GetcMovNro(lsMovNroAux)
                    Next
            End If
        End If
        'end PASI
        
        'PASI20150918 ERS0472015
        If lbModifica Then
            bestado = False
            For nx = 1 To fgDetalle.Rows - 1
                If Not (oMov.obtieneExisteItemComprobanteOrden(CInt(fgDetalle.TextMatrix(nx, 0)), CLng(nMovNroOrdAnt))) = 0 Then
                    bestado = True
                    Exit For
                End If
            Next
            If bestado Then
                nMovNroComp = oMov.ObtieneMovComprobantexOrden(CLng(nMovNroOrdAnt))
                oMov.InsertaMovComprobanteOrden nMovNroComp, lnMovNro
            End If
        End If
        'END PASI
        
        oMov.CommitTrans
        lTransActiva = False
        OK = True
        Dim lsImpre As String
   
    If lOrdenCompra And Not lbModifica And Not lbImprime Then
        
           If MsgBox(" ¿ Desea Ingresar nuevo Documento ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbYes Then
              ReDim rsCosto(100)
              gcDocNro = oMov.GeneraDocNro(Val(gcDocTpo), Right(Me.cmbmoneda.Text, 1), Year(gdFecSis), "OSC") 'NAGL 20191212 Agregó "OSC"
              For n = 1 To fgDetalle.Rows - 1
                  For nItem = 1 To fgDetalle.Cols - 1
                      fgDetalle.TextMatrix(n, nItem) = ""
                  Next
              Next
              fgObj.Rows = 2
              For nItem = 1 To fgObj.Cols - 1
                  fgObj.TextMatrix(1, nItem) = ""
              Next
              txtTot = ""
              txtMovDesc = ""
              LblDoc.Caption = UCase(sDocDesc) & "    Nº " & gcDocNro
              txtDiasAtraso.Text = ""
              txtCotNro.Text = ""
              txtPersona.Text = ""
              txtProvNom.Text = ""
              txtArea.Text = ""
              txtAgeDesc.Text = ""
              txtCodLugarEntrega.Text = ""
              txtLugarEntrega.Text = ""
              lblDocOCD.Visible = False
              fgDetalle.SetFocus
           Else
              Unload Me
           End If
        Else
           Unload Me
        End If
'    End If
    CmdAceptar.Enabled = True
    
    Exit Sub
ErrAceptar:
    If lTransActiva Then
        oMov.RollbackTrans
    End If
    lTransActiva = False
    CmdAceptar.Enabled = True
    MsgBox TextErr(Err.Description), vbCritical, "Error de Actualización"
End Sub

Public Property Get lOk() As Boolean
    lOk = OK
End Property

Public Property Let lOk(ByVal vNewValue As Boolean)
OK = vNewValue
End Property

Private Function ValidaAgencia(sAgeCod As String) As Boolean
'ValidaAgencia = False
'sSQL = "SELECT h.cObjetoCod, h.cObjetoDesc, " _
'   & "  h.nObjetoNiv FROM   " & gcCentralCom & "Objeto h WHERE  " _
'   & "  cObjetoCod LIKE '11%' and cObjetoCod like '" & IIf(sAgeCod = "", "11", sAgeCod & "%' ")
'Set rs = CargaRecord(sSQL)
'If rs.EOF Then
'   MsgBox "Area funcional no existe", vbInformation, "¡Aviso!"
'Else
'   ValidaAgencia = True
'End If
End Function

Private Function ValidaDatos() As Boolean
    Dim nx As Integer 'PASI20150917 ERS0472015
    Dim rsItemDet As ADODB.Recordset
    Dim bMod As Boolean
    Dim m, n As Integer
    Dim olog As DLogGeneral
    Set olog = New DLogGeneral
    Dim lsObjetos() As String
    Dim lnValorCelda As Currency
    Dim lsObjetosCont As String
    Dim lnCantFilas As Integer
    'end PASI
    ValidaDatos = False
    If txtTot = "" Then
       MsgBox "Monto de Documento debe ser mayor a Cero", vbInformation, "¡Aviso!"
       fgDetalle.SetFocus
       Exit Function
    End If
    If nVal(txtTot) <= 0 Then
       MsgBox "Monto de Documento debe ser mayor a Cero", vbInformation, "¡Aviso!"
       fgDetalle.SetFocus
       Exit Function
    End If
    If ValidaFecha(txtFecha) <> "" Then
       MsgBox "Fecha de documento no válida", vbInformation, "¡Aviso!"
       txtFecha.SetFocus
       Exit Function
    End If
    If lsCodPers = "" Then
       MsgBox "Falta Datos del Proveedor", vbInformation, "¡Aviso!"
       txtPersona.SetFocus
       Exit Function
    End If
    If txtMovDesc = "" Then
       MsgBox "Falta ingresar Observaciones del documento", vbInformation, "¡Aviso!"
       txtMovDesc.SetFocus
       Exit Function
    End If
    If cboFormaPago.ListIndex = -1 Then
       MsgBox "Falta Indicar la Forma de Pago", vbInformation, "¡Aviso!"
       cboFormaPago.SetFocus
       Exit Function
    End If
    
'    If txtLugarEntrega = "" And lbBienes Then
'       MsgBox "Falta indicar el Lugar de Entrega de la Orden", vbInformation, "¡Aviso!"
'       txtCodLugarEntrega.SetFocus
'       Exit Function
'    End If
    
    If Me.cmbmoneda.Text = "" Then
       MsgBox "Falta indicar la moneda", vbInformation, "¡Aviso!"
       cmbmoneda.SetFocus
       Exit Function
    End If
    'PASI20140818 ERS0772014
    If chkContSuministro.value = 1 Then
              
        If cboContSuministro.Text = "" Then
            MsgBox "Falta Seleccionar el Contrato de Suministro", "¡Aviso!"
            cboContSuministro.SetFocus
            Exit Function
        End If
        For n = 1 To fgDetalle.Rows - 1
            If fgDetalle.TextMatrix(n, 1) <> "" And fgDetalle.TextMatrix(n, 2) <> "" Then
                If Not olog.ComparaBSContrato(Trim(Right(cboContSuministro.Text, 100)), fgDetalle.TextMatrix(n, 2)) Then
                    MsgBox "Los Bienes/Servicios Ingresados no Pertenecen al Contrato", vbInformation, "¡Aviso!"
                    fgDetalle.SetFocus
                    Exit Function
                End If
            End If
        Next
        Dim lsMontoContrato() As Currency
        Dim k As Integer
        Dim bestado As Boolean
        If sProvCod = "" Then
            ReDim Preserve lsMontoContrato(2, 0)
            For n = 1 To fgDetalle.Rows - 1
                If fgDetalle.TextMatrix(n, 1) <> "" And fgDetalle.TextMatrix(n, 2) <> "" Then
                    If n = 1 Then
                        ReDim Preserve lsMontoContrato(2, 1)
                        lsMontoContrato(1, 1) = fgDetalle.TextMatrix(n, 2)
                        lsMontoContrato(2, 1) = fgDetalle.TextMatrix(n, 8)
                    Else
                        For k = 1 To UBound(lsMontoContrato, 2)
                            If fgDetalle.TextMatrix(n, 2) = lsMontoContrato(1, k) Then
                                lsMontoContrato(2, k) = lsMontoContrato(2, k) + fgDetalle.TextMatrix(n, 8)
                                bestado = True
                            End If
                        Next
                        If Not bestado Then
                            ReDim Preserve lsMontoContrato(2, UBound(lsMontoContrato, 2) + 1)
                            lsMontoContrato(1, UBound(lsMontoContrato, 2)) = fgDetalle.TextMatrix(n, 2)
                            lsMontoContrato(2, UBound(lsMontoContrato, 2)) = fgDetalle.TextMatrix(n, 8)
                        End If
                        bestado = False
                    End If
                End If
            Next
            For k = 1 To UBound(lsMontoContrato, 2)
                If Not olog.ComparaSaldoBSContrato(Trim(Right(cboContSuministro.Text, 100)), lsMontoContrato(1, k), lsMontoContrato(2, k)) Then
                    MsgBox "El contrato no tiene saldo suficiente para la solicitud realizada", vbInformation, "¡Aviso!"
                    fgDetalle.SetFocus
                    Exit Function
                End If
            Next
        Else
        Dim oMov As DMov
        Set oMov = New DMov
            ReDim Preserve lsMontoContrato(2, 0)
            For n = 1 To fgDetalle.Rows - 1
                If n = 1 Then
                    ReDim Preserve lsMontoContrato(2, 1)
                    lsMontoContrato(1, 1) = fgDetalle.TextMatrix(n, 2)
                    lsMontoContrato(2, 1) = fgDetalle.TextMatrix(n, 8)
                Else
                    For k = 1 To UBound(lsMontoContrato, 2)
                        If fgDetalle.TextMatrix(n, 2) = lsMontoContrato(1, k) Then
                            lsMontoContrato(2, k) = lsMontoContrato(2, k) + fgDetalle.TextMatrix(n, 8)
                            bestado = True
                        End If
                    Next
                    If Not bestado Then
                        ReDim Preserve lsMontoContrato(2, UBound(lsMontoContrato, 2) + 1)
                        lsMontoContrato(1, UBound(lsMontoContrato, 2)) = fgDetalle.TextMatrix(n, 2)
                        lsMontoContrato(2, UBound(lsMontoContrato, 2)) = fgDetalle.TextMatrix(n, 8)
                    End If
                    bestado = False
                End If
            Next
            For k = 1 To UBound(lsMontoContrato, 2)
                If Not olog.ComparaSaldoBSEditContrato(Trim(Right(cboContSuministro.Text, 100)), lsMontoContrato(1, k), lsMontoContrato(2, k), oMov.GetcMovNro(gcMovNro)) Then
                    MsgBox "El contrato no tiene saldo suficiente para la solicitud realizada", vbInformation, "¡Aviso!"
                    fgDetalle.SetFocus
                    Exit Function
                End If
            Next
        End If
    End If
    'end PASI
'PASI20150917 ERS0472015
'Validar los Items ya utilizados
If lbModifica Then
    For nx = 1 To fgDetalle.Rows - 1
        If Not (olog.obtieneExisteItemComprobanteOrden(CInt(fgDetalle.TextMatrix(nx, 0)), CLng(gcMovNro))) = 0 Then
            bMod = False
            If fgDetalle.TextMatrix(nx, 11) = "B" Then
                Set rsItemDet = olog.ObtieneItemDetxOC(CLng(gcMovNro), fgDetalle.TextMatrix(nx, 0))
                If Not rsItemDet.EOF Then
                    If fgDetalle.TextMatrix(nx, 1) <> rsItemDet!cAgeCod Then bMod = True
                    If fgDetalle.TextMatrix(nx, 2) <> rsItemDet!cBSCod Then bMod = True
                    If fgDetalle.TextMatrix(nx, 3) <> rsItemDet!cDescrip Then bMod = True
                    If CLng(fgDetalle.TextMatrix(nx, 8)) <> CLng(Format(rsItemDet!nMovImporte, gsFormatoNumeroView)) Then bMod = True
                End If
            Else
                Set rsItemDet = olog.ObtieneItemDetxOS(CLng(gcMovNro), fgDetalle.TextMatrix(nx, 0))
                If Not rsItemDet.EOF Then
                    If fgDetalle.TextMatrix(nx, 1) <> rsItemDet!cAgeCod Then bMod = True
                    If fgDetalle.TextMatrix(nx, 9) <> rsItemDet!cCtaContCod Then bMod = True
                    If fgDetalle.TextMatrix(nx, 3) <> rsItemDet!cDescrip Then bMod = True
                    If CLng(fgDetalle.TextMatrix(nx, 8)) <> CLng(Format(rsItemDet!nMovImporte, gsFormatoNumeroView)) Then bMod = True
                End If
            End If
            If bMod Then
                MsgBox "No se pueden guardar los cambios por que al parecer existen items que no pueden ser modificados. ", vbInformation, "¡Aviso!"
                fgDetalle.SetFocus
                Exit Function
            End If
        End If
    Next
End If
'END PASI
    ValidaDatos = True
End Function

Private Function ValidaCta() As Boolean
    Dim oCon As DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
    ValidaCta = False
    If Len(txtObj) = 0 Then
       txtObj.Visible = False
       cmdExaminar.Visible = False
       Exit Function
    End If
    sSql = "SELECT a.cCtaContCod, b.cCtaContDesc, a.cOpeCtaDH " _
         & "FROM  " & gcCentralCom & "OpeCta a,  " & gcCentralCom & "CtaCont b " _
         & "WHERE b.cCtaContCod = a.cCtaContCod AND ((a.cOpeCod='" & gcOpeCod & "') AND (a.cOpeCtaDH='D')) " _
         & "and a.cCtaContCod = '" & txtObj & "'"
    Set rs = oCon.CargaRecordSet(sSql)
    If Not rs.EOF Then
        If fgDetalle.TextMatrix(fgDetalle.row, 11) = "S" Then
            AsignaObjetosSer rs!cCtaContCod, rs!cCtaContDesc
        Else
            AsignaObjetos rs!cCtaContCod, rs!cCtaContDesc
        End If
    Else
       MsgBox "Objeto no encontrado...!", vbCritical, "Error de Búsqueda"
       Exit Function
    End If
    ValidaCta = True
End Function

Private Sub FormatoObjeto()
fgObj.Cols = 6
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
fgObj.ColWidth(5) = 780
fgObj.ColAlignment(1) = 7
fgObj.ColAlignment(2) = 1
End Sub

Private Sub FormatoOrden()
fgDetalle.TextMatrix(0, 0) = "#"
fgDetalle.TextMatrix(0, 2) = "Objeto"
fgDetalle.TextMatrix(0, 3) = "Descripción"
fgDetalle.TextMatrix(0, 4) = "Unidad"
fgDetalle.TextMatrix(0, 5) = "Solicitado"
fgDetalle.TextMatrix(0, 6) = "P.Unitario"
fgDetalle.TextMatrix(0, 7) = "Saldo"
fgDetalle.TextMatrix(0, 8) = "Sub Total"
fgDetalle.TextMatrix(0, 12) = "Plazo"
'ALPA 20091110****************************************************
fgDetalle.TextMatrix(0, 1) = "Age.Destino"
'*****************************************************************
fgDetalle.ColWidth(0) = 335
fgDetalle.ColWidth(2) = 1500
fgDetalle.ColWidth(3) = 3500
fgDetalle.ColWidth(4) = 880
fgDetalle.ColWidth(5) = 1200
fgDetalle.ColWidth(6) = 1200
fgDetalle.ColWidth(7) = 0
fgDetalle.ColWidth(8) = 1200
fgDetalle.ColWidth(9) = 0
fgDetalle.ColWidth(10) = 0
fgDetalle.ColWidth(11) = 0
fgDetalle.ColWidth(12) = 0
'ALPA 20091110****************************************************
fgDetalle.ColWidth(1) = 1200
'*****************************************************************
fgDetalle.ColAlignment(2) = 1
fgDetalle.ColAlignmentFixed(0) = 4
fgDetalle.ColAlignmentFixed(5) = 7
fgDetalle.ColAlignmentFixed(6) = 7
fgDetalle.ColAlignmentFixed(7) = 7
fgDetalle.ColAlignmentFixed(8) = 7
fgDetalle.ColAlignmentFixed(12) = 4
fgDetalle.RowHeight(-1) = 285
End Sub

Private Sub AsignaObjetos(sCtaCod As String, sCtaDes As String)
    Dim rsObj As ADODB.Recordset
    Set rsObj = New ADODB.Recordset
    Dim nNiv As Integer
    Dim nObj As Integer
    Dim nObjs As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If fgDetalle.TextMatrix(fgDetalle.row, 0) <> "" Then
        EliminaFgObj fgDetalle.TextMatrix(fgDetalle.row, 0)
    End If
    fgDetalle = sCtaCod
    fgDetalle.TextMatrix(fgDetalle.row, 3) = sCtaDes
    fgDetalle.TextMatrix(fgDetalle.row, 9) = sCtaCod
    
    oCon.AbreConexion
    sSql = "SELECT MAX(nCtaObjOrden) as nNiveles FROM CtaObj WHERE cCtaContCod = '" & sCtaCod & "' and cObjetoCod <> '00' "
    Set rs = oCon.CargaRecordSet(sSql)
    
    nObjs = IIf(IsNull(rs!nNiveles), 0, rs!nNiveles)
    For nObj = 1 To nObjs
       sSql = "SELECT co.cObjetoCod, co.nCtaObjOrden,  4 nObjetoNiv, co.nCtaObjNiv, co.cCtaObjFiltro, co.cCtaObjImpre FROM CtaBS CO JOIN BienesServicios o ON o.cBSCod like co.cObjetoCod  WHERE co.cCtaContCod = '" & sCtaCod & "' and co.nCtaObjOrden = '" & nObj & "'"
       Set rs = oCon.CargaRecordSet(sSql)
       AceptaOK = False
       If rs!cObjetoCod <> "00" Then
          AdicionaObj fgDetalle, fgDetalle.TextMatrix(fgDetalle.row, 0), rs
          nNiv = rs!nObjetoNiv + rs!nCtaObjNiv
          If rs.RecordCount = 1 Then 'Una Cuenta por Orden
             sSql = " " & gcCentralCom & "spGetTreeObj '" & rs!cObjetoCod & "', " & nNiv & ",'" & rs!cCtaObjFiltro & "'"
          Else  'Para varios Objetos en una misma Orden asignado a la Cuenta
             sSql = " SELECT o.* FROM BienesServicios O JOIN CtaBS co ON co.cObjetoCod = substring(o.cBSCod,1,LEN(co.cObjetoCod)) WHERE co.cCtaContCod = '" & sCtaCod & "' and co.nCtaObjOrden = '" & nObj & "' and Exists (select * from BienesServicios as a where a.cBSCod like o.cBSCod + '%') order by o.cBSCod"
          End If
          sSql = "SELECT *, Len(cBSCod) - 2 nObjetoNiv FROM bienesServicios bs WHERE cBSCod Like '" & rs!cObjetoCod & "'+'%' and Exists (select * from bienesservicios as a where a.cBSCod like bs.cBSCod+'%') order by bs.cbscod"
          If rsObj.State = adStateOpen Then rsObj.Close: Set rsObj = Nothing
          Set rsObj = oCon.CargaRecordSet(sSql)
          If Not rsObj.EOF Then
             frmDescObjeto.Inicio rsObj, "", 7
             If rsObj.State = adStateOpen Then rsObj.Close: Set rsObj = Nothing
             If frmDescObjeto.lOk Then
                AceptaOK = True
             End If
          End If
       End If
       If AceptaOK Then
          fgObj.TextMatrix(fgObj.row, 2) = gaObj(0, 0, UBound(gaObj, 3) - 1) 'CodObj
          fgObj.TextMatrix(fgObj.row, 3) = gaObj(0, 1, UBound(gaObj, 3) - 1) 'DesObj
          sSql = "SELECT cCtaObjMascara FROM " & gcCentralCom & "CtaObjFiltro WHERE cCtaContCod = '" & fgDetalle & "' and cObjetoCod = '" & fgObj.TextMatrix(fgObj.row, 2) & "'"
          If rsObj.State = adStateOpen Then rsObj.Close: Set rsObj = Nothing
            Set rsObj = oCon.CargaRecordSet(sSql)
          If rsObj.EOF Then
             fgObj.TextMatrix(fgObj.row, 4) = ""
          Else
             fgObj.TextMatrix(fgObj.row, 4) = rsObj!cCtaObjMascara
          End If
       Else
          If rs!cObjetoCod <> "00" Then
             If fgDetalle.TextMatrix(fgDetalle.row, 0) <> "" Then
                EliminaObjeto fgDetalle.TextMatrix(fgDetalle.row, 0)
            End If
             If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
             If fgDetalle.Enabled Then
                fgDetalle.SetFocus
             End If
             Exit Sub
          End If
       End If
       If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
    Next
End Sub

Private Sub EliminaFgObj(nItem As Integer)
Dim k  As Integer, m As Integer
k = 1
Do While k < fgObj.Rows
   If Len(fgObj.TextMatrix(k, 1)) > 0 Then
      If Val(fgObj.TextMatrix(k, 0)) = nItem Then
         EliminaRow fgObj, k
         k = k - 1
      Else
         k = k + 1
      End If
   Else
      k = k + 1
   End If
Loop
End Sub

Private Sub AdicionaObj(sCodCta As String, nFila As Integer, rs As ADODB.Recordset)
   Dim nItem As Integer
   AdicionaRow fgObj
   nItem = fgObj.row
   fgObj.TextMatrix(nItem, 0) = nFila
   fgObj.TextMatrix(nItem, 1) = rs!nCtaObjOrden
End Sub

Private Sub EliminaObjeto(nItem As Integer)
    EliminaFgObj nItem
    If Len(fgDetalle.TextMatrix(1, 1)) > 0 Then
       RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.row, 0)
    End If
End Sub

Private Sub RefrescaFgObj(nItem As Integer)
    Dim k  As Integer
    For k = 1 To fgObj.Rows - 1
        If Len(fgObj.TextMatrix(k, 1)) Then
           If fgObj.TextMatrix(k, 0) = nItem Then
              fgObj.RowHeight(k) = 285
           Else
              fgObj.RowHeight(k) = 0
           End If
        End If
    Next
End Sub

Private Sub AsignaObjetosSer(sCtaCod As String, sCtaDes As String)
    Dim rsObj As ADODB.Recordset
    Set rsObj = New ADODB.Recordset
    Dim nNiv As Integer
    Dim nObj As Integer
    Dim nObjs As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oCtaCont As DCtaCont
    Set oCtaCont = New DCtaCont
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    Dim oRHAreas As DActualizaDatosArea
    Dim oCtaIf As NCajaCtaIF
    Dim oEfect As Defectivo
    Set oRHAreas = New DActualizaDatosArea
    Set oCtaIf = New NCajaCtaIF
    Set oEfect = New Defectivo
    Dim oDescObj As ClassDescObjeto
    Set oDescObj = New ClassDescObjeto
    Dim oContFunct As NContFunciones
    Set oContFunct = New NContFunciones
    Dim lsRaiz  As String
    Dim lsFiltro  As String
    
    
    oDescObj.lbUltNivel = True
    oCon.AbreConexion
      EliminaFgObj fgDetalle.TextMatrix(fgDetalle.row, 0)
      fgDetalle = sCtaCod
      fgDetalle.TextMatrix(fgDetalle.row, 3) = sCtaDes
      fgDetalle.TextMatrix(fgDetalle.row, 9) = sCtaCod
      sSql = "SELECT MAX(nCtaObjOrden) as nNiveles FROM CtaObj WHERE cCtaContCod = '" & sCtaCod & "' and cObjetoCod <> '00' "
      Set rs = oCon.CargaRecordSet(sSql)
      nObjs = IIf(IsNull(rs!nNiveles), 0, rs!nNiveles)
      
    Set rs1 = oCtaCont.CargaCtaObj(sCtaCod, , True)
    If Not rs1.EOF And Not rs1.BOF Then
        Do While Not rs1.EOF
            lsRaiz = ""
            lsFiltro = ""
            Set rs = New ADODB.Recordset
            Select Case Val(rs1!cObjetoCod)
                Case ObjCMACAgencias
                    Set rs = oRHAreas.GetAgencias()
                Case ObjCMACAgenciaArea
                    lsRaiz = "Unidades Organizacionales"
                    Set rs = oRHAreas.GetAgenciasAreas()
                Case ObjCMACArea
                    Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
                Case ObjEntidadesFinancieras
                    lsRaiz = "Cuentas de Entidades Financieras"
                    Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, sCtaCod)
                Case ObjDescomEfectivo
                    Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
                Case ObjPersona
                    Set rs = Nothing
                Case Else
                    lsRaiz = "Varios"
                    Set rs = GetObjetos(rs1!cObjetoCod)
            End Select
            If Not rs Is Nothing Then
                If rs.State = adStateOpen Then
                    If Not rs.EOF And Not rs.BOF Then
                        If rs.RecordCount > 1 Then
                            oDescObj.Show rs, "", lsRaiz
                            If oDescObj.lbOK Then
                                'lsFiltro = oContFunct.GetFiltroObjetos(Val(rs1!cObjetoCod), sCtaCod, oDescObj.gsSelecCod, False)
                                lsFiltro = oContFunct.GetFiltroObjetos(Trim(rs1!cObjetoCod), sCtaCod, oDescObj.gsSelecCod, False)
                                AdicionaObj sCtaCod, fgDetalle.TextMatrix(fgDetalle.row, 0), rs1   '!cCtaObjOrden, oDescObj.gsSelecCod, _
                                            oDescObj.gsSelecDesc, lsFiltro, rs1!cObjetoCod
                                fgObj.TextMatrix(fgObj.row, 2) = oDescObj.gsSelecCod
                                fgObj.TextMatrix(fgObj.row, 3) = oDescObj.gsSelecDesc
                                fgObj.TextMatrix(fgObj.row, 4) = lsFiltro
                                fgObj.TextMatrix(fgObj.row, 5) = rs1!cObjetoCod
                            Else
                                EliminaFgObj fgDetalle.TextMatrix(fgDetalle.row, 0)
                                'fgDetalle.EliminaFila fgDetalle.Row, False
                                Exit Do
                            End If
                        Else
                            AdicionaObj sCtaCod, fgDetalle.TextMatrix(fgDetalle.row, 0), rs1 '!cCtaObjOrden, rs1!cObjetoCod, _
                                            rs1!cObjetoDesc, lsFiltro, rs1!cObjetoCod
                        End If
                    End If
                End If
            End If
            rs1.MoveNext
        Loop
    End If
    rs1.Close
    Set rs1 = Nothing
    Set oDescObj = Nothing
    
    Set oCtaCont = Nothing
    Set oCtaIf = Nothing
    Set oEfect = Nothing
      
    Exit Sub
      For nObj = 1 To nObjs
         sSql = "SELECT co.cObjetoCod, co.nCtaObjOrden, o.nObjetoNiv, co.nCtaObjNiv, co.cCtaObjFiltro, co.cCtaObjImpre FROM " & gcCentralCom & "CtaObj co JOIN " & gcCentralCom & "Objeto o ON o.cObjetoCod = co.cObjetoCod WHERE co.cCtaContCod = '" & sCtaCod & "' and co.nCtaObjOrden = '" & nObj & "'"
         Set rs = oCon.CargaRecordSet(sSql)
         AceptaOK = False
         If rs!cObjetoCod <> "00" Then
            AdicionaObj fgDetalle, fgDetalle.TextMatrix(fgDetalle.row, 0), rs
            nNiv = rs!nObjetoNiv + rs!nCtaObjNiv
            If rs.RecordCount = 1 Then 'Una Cuenta por Orden
               sSql = " spGetTreeObj '" & rs!cObjetoCod & "', " & nNiv & ",'" & rs!cCtaObjFiltro & "'"
            Else  'Para varios Objetos en una misma Orden asignado a la Cuenta
               sSql = "SELECT o.* FROM " & gcCentralCom & "Objeto O JOIN " & gcCentralCom & "CtaObj co ON co.cObjetoCod = substring(o.cObjetoCod,1,LEN(co.cObjetoCod)) WHERE co.cCtaContCod = '" & sCtaCod & "' and co.cCtaObjOrden = '" & nObj & "' and Exists (select * from " & gcCentralCom & "Objeto as a where a.cobjetocod like o.cobjetocod+'%' and a.nObjetoNiv = " & nNiv & ") order by o.cobjetocod "
            End If
            If rsObj.State = adStateOpen Then rsObj.Close: Set rsObj = Nothing
            Set rsObj = oCon.CargaRecordSet(sSql)
            
            If Not rsObj.EOF Then
               frmDescObjeto.Inicio rsObj, "", nNiv
               If rsObj.State = adStateOpen Then rsObj.Close: Set rsObj = Nothing
               If frmDescObjeto.lOk Then
                  AceptaOK = True
               End If
            End If
         End If
         If AceptaOK Then
            fgObj.TextMatrix(fgObj.row, 2) = gaObj(0, 0, UBound(gaObj, 3) - 1) 'CodObj
            fgObj.TextMatrix(fgObj.row, 3) = gaObj(0, 1, UBound(gaObj, 3) - 1) 'DesObj
            sSql = "SELECT cCtaObjMascara FROM " & gcCentralCom & "CtaObjFiltro WHERE cCtaContCod = '" & fgDetalle & "' and cObjetoCod = '" & fgObj.TextMatrix(fgObj.row, 2) & "'"
            If rsObj.State = adStateOpen Then rsObj.Close: Set rsObj = Nothing
            Set rsObj = oCon.CargaRecordSet(sSql)
            
            If rsObj.EOF Then
               fgObj.TextMatrix(fgObj.row, 4) = ""
            Else
               fgObj.TextMatrix(fgObj.row, 4) = rsObj!cCtaObjMascara
            End If
         Else
            If rs!cObjetoCod <> "00" Then
               EliminaObjeto fgDetalle.TextMatrix(fgDetalle.row, 0)
               If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
               If fgDetalle.Enabled Then
                  fgDetalle.SetFocus
               End If
               Exit Sub
            End If
         End If
         If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
      Next
End Sub

Private Function GetFechaMov(cMovNro, lDia As Boolean) As String
    Dim lFec As Date
    lFec = Mid(cMovNro, 7, 2) & "/" & Mid(cMovNro, 5, 2) & "/" & Mid(cMovNro, 1, 4)
    If lDia Then
       GetFechaMov = Format(lFec, gsFormatoFechaView)
    Else
       GetFechaMov = Format(lFec, gsFormatoFecha)
    End If
End Function

Sub emiteProveedor()
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oProv As DLogProveedor
    Set oProv = New DLogProveedor
    Me.txtProvNom.Text = txtPersona.psDescripcion
    lsCodPers = txtPersona.psCodigoPersona
    If txtPersona.psDescripcion <> "" Then
        Set rs = oProv.GetProveedorAgeRetBuenCont(lsCodPers)
        If rs.EOF And rs.BOF Then
            'MsgBox "La persona ingresada no esta registrada como proveedor o tiene el estado de Desactivado, debe regsitrarlo o activarlo.", vbInformation, "Aviso"
        Else
            Me.chkBuneCOnt.value = IIf(rs.Fields(1), 1, 0)
            Me.chkRetencion.value = IIf(rs.Fields(0), 1, 0)
        End If
        Me.txtArea.SetFocus
    End If
    
End Sub
'PASI20140724 TI-ERS077-2014
Private Sub chkContSuministro_Click()
    If chkContSuministro.value = 1 Then
        cboContSuministro.Enabled = True
    Else
        cboContSuministro.Enabled = False
    End If
End Sub
'end PASI
Private Sub CargaComboSuministro() 'PASIERS0772014
     Dim lsNContrato As String 'PASI20140724 TI-ERS077-2014
    Dim olog As DLogGeneral 'PASI20140724 TI-ERS077-2014
    Set olog = New DLogGeneral
    Dim rsComboSum As ADODB.Recordset 'PASI20140724 TI-ERS077-2014
    Set rsComboSum = New ADODB.Recordset
    
    Set rsComboSum = olog.ComboContratoSuministro("4,5", Mid(gcOpeCod, 3, 1))
    RSLlenaCombo rsComboSum, Me.cboContSuministro
    cboContSuministro.Enabled = False
    
    If gcOpeCod = "501207" Or gcOpeCod = "502207" Then
        cmdSeleccion.Visible = False
        chkContSuministro.Visible = False 'PASI20140930 ERS0772014
        cboContSuministro.Visible = False 'PASI20140930 ERS0772014
    End If
    
End Sub
'PASI20140724 ERS0772014
Public Sub RSLlenaCombo(prs As ADODB.Recordset, psCombo As ComboBox, Optional pnPosCod As Integer = 0, Optional pnPosDes As Integer = 1, Optional pbPresentaCodigo As Boolean = True)
    If Not prs Is Nothing Then
        If Not prs.EOF Then
            psCombo.Clear
            Do While Not prs.EOF
                If pbPresentaCodigo Then
                    psCombo.AddItem Trim(prs(pnPosDes)) & Space(100) & Trim(prs(pnPosCod))
                Else
                    psCombo.AddItem Trim(prs(pnPosCod)) & "  " & Trim(prs(pnPosDes))
                End If
                prs.MoveNext
            Loop
        End If
    End If
End Sub
'END PASI

