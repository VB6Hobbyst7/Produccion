VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmLogOCompra 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6810
   ClientLeft      =   930
   ClientTop       =   1260
   ClientWidth     =   11055
   Icon            =   "frmLogOCompra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11055
   Visible         =   0   'False
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
      TabIndex        =   65
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
      TabIndex        =   64
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
      TabIndex        =   53
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
      TabIndex        =   51
      Top             =   15
      Width           =   1740
      Begin VB.ComboBox cmbMoneda 
         Enabled         =   0   'False
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   52
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
      Left            =   120
      TabIndex        =   41
      Top             =   2310
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
      Begin SICMACT.TxtBuscar txtCodLugarEntrega 
         Height          =   315
         Left            =   135
         TabIndex        =   48
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
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
      MaxLength       =   1000
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
      Top             =   6150
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
      Left            =   8670
      TabIndex        =   17
      Top             =   6150
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
      Begin SICMACT.TxtBuscar txtPersona 
         Height          =   315
         Left            =   135
         TabIndex        =   50
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
         TabIndex        =   54
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
            TabIndex        =   56
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
            TabIndex        =   55
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
      Left            =   9840
      TabIndex        =   18
      Top             =   6150
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
      Cols            =   12
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
      _Band(0).Cols   =   12
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
      Begin SICMACT.TxtBuscar txtArea 
         Height          =   315
         Left            =   120
         TabIndex        =   49
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
      Left            =   4320
      TabIndex        =   30
      Top             =   2310
      Width           =   6675
      Begin VB.ComboBox cmbTipoOC 
         Height          =   315
         ItemData        =   "frmLogOCompra.frx":0724
         Left            =   5400
         List            =   "frmLogOCompra.frx":072E
         Style           =   2  'Dropdown List
         TabIndex        =   59
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
         Format          =   63242241
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
         TabIndex        =   58
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
         TabIndex        =   62
         Top             =   150
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtDiasAtraso 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2025
         MaxLength       =   40
         TabIndex        =   60
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
         TabIndex        =   63
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
         TabIndex        =   61
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.CommandButton cmdDocumento 
      Caption         =   "&Imprimir"
      Height          =   315
      Left            =   8700
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
      TabIndex        =   57
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
Dim sCtaCod    As String, sCtaDesc As String
Dim sProvCod   As String
Dim sCtaProvis As String
Dim sBS        As String
Dim lNewProv   As Boolean, lbBienes As Boolean, lbModifica As Boolean
Dim sDocDesc   As String
Dim lnColorBien As Double
Dim lnColorServ As Double
Dim lsMovNro    As String
Dim lbImprime   As Boolean
Dim lsEstado    As String
Dim AceptaOK As Boolean
Dim lsCodPers As String

Dim lsDocNroOCD As String

Public Sub Inicio(LlenaObj As Boolean, pcOpeCod As String, Optional ProvCod As String = "", Optional lBienes As Boolean = True, Optional pbModifica As Boolean = False, Optional pbImprime As Boolean = False, Optional psEstado As String = "", Optional psCaption As String = "")
    lLlenaObj = LlenaObj
    sProvCod = ProvCod
    lsCodPers = ProvCod
    lbBienes = lBienes
    lbModifica = pbModifica
    lbImprime = pbImprime
    lsEstado = psEstado
    gcOpeCod = pcOpeCod
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
   Dim sPrenombre As String
   
   Dim ofun As NContFunciones
   Set ofun = New NContFunciones
   
   If lbBienes = True Then
      If Right(cmbTipoOC.Text, 1) = "D" Then
        sTipodoc = "130"
        sPrenombre = " DIRECTA"
      ElseIf Right(cmbTipoOC.Text, 1) = "P" Then
        sTipodoc = "133"
        sPrenombre = " PROCESO"
      End If
   Else
        If Right(cmbTipoOC.Text, 1) = "D" Then
        sTipodoc = "132"
        sPrenombre = " DIRECTA"
        ElseIf Right(cmbTipoOC.Text, 1) = "P" Then
        sTipodoc = "134"
        sPrenombre = " PROCESO"
        End If
   End If
   
   If Right(cmbTipoOC.Text, 1) = "D" Or Right(cmbTipoOC.Text, 1) = "P" Then
   
   lblDocOCD.Visible = True
   
   If lbModifica = True Then
      
       
        sSql = " select  cDocNro from movdoc where nMovNro = " & gcMovNro & " and nDocTpo ='" & sTipodoc & "' "
        If oCon.AbreConexion = True Then
                Set rs = oCon.CargaRecordSet(sSql)
                If rs.EOF Then
                     lblDocOCD.Caption = ""
                    lblDocOCD.Visible = False
                    Else
                    lblDocOCD.Caption = UCase(sDocDesc) & sPrenombre & " Nº " & rs!cDocNro
                End If
        End If
   Else
    
        lsDocNroOCD = ofun.GeneraDocNro(CInt(sTipodoc), Mid(gcOpeCod, 3, 1), Year(gdFecSis))
        lblDocOCD.Caption = UCase(sDocDesc) & sPrenombre & " Nº " & lsDocNroOCD
   End If
   
   
   Else
   lblDocOCD.Visible = False
End If

End Sub

Private Sub cmdServicio_Click()
If fgDetalle.TextMatrix(fgDetalle.Row, 10) = "B" Then
   If fgDetalle.TextMatrix(fgDetalle.Row, 1) <> "" Then
      MsgBox "Item asignado para Ingreso de Bienes y ya tiene un Valor Asignado", vbInformation, "¡Aviso!"
      Exit Sub
   End If
   fgDetalle.TextMatrix(fgDetalle.Row, 10) = "S"
   FlexBackColor fgDetalle, fgDetalle.Row, lnColorServ
   cmdServicio.Caption = "&Bienes"
Else
   If fgDetalle.TextMatrix(fgDetalle.Row, 1) <> "" Then
      MsgBox "Item asignado para Ingreso de Servicios y ya tiene un Valor Asignado", vbInformation, "¡Aviso!"
      Exit Sub
   End If
   fgDetalle.TextMatrix(fgDetalle.Row, 10) = "B"
   FlexBackColor fgDetalle, fgDetalle.Row, lnColorBien
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
If fgDetalle.TextMatrix(fgDetalle.Row, 10) = "B" Then
   cmdServicio.Caption = "&Servicio"
Else
   cmdServicio.Caption = "&Bienes"
End If
If fgDetalle.TextMatrix(fgDetalle.Row, 0) <> "" Then
    RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.Row, 0)
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
   If Val(Format(fgDetalle.TextMatrix(fgDetalle.Rows - 1, 7), gcFormDato)) <> 0 And _
      Len(Format(fgDetalle.TextMatrix(fgDetalle.Rows - 1, 1), gcFormDato)) > 0 Then
      AdicionaRow fgDetalle, fgDetalle.Rows
      If lbBienes Then
         fgDetalle.TextMatrix(fgDetalle.Row, 10) = "B"
      Else
         fgDetalle.TextMatrix(fgDetalle.Row, 10) = "S"
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
nSalto = IIf(fgDetalle.Row < fgDetalle.RowSel, 1, -1)
For nPos = fgDetalle.Row To fgDetalle.RowSel Step nSalto
    fgDetalle.TextMatrix(nPos, 5) = fgDetalle.TextMatrix(nPos, 4)
Next
SumasDoc
End Sub

Private Sub mnuEliminar_Click()
If fgDetalle.TextMatrix(fgDetalle.Row, 0) <> "" Then
   EliminaCuenta fgDetalle.TextMatrix(fgDetalle.Row, 1), fgDetalle.TextMatrix(fgDetalle.Row, 0)
   SumasDoc
   If fgDetalle.Enabled Then
      fgDetalle.SetFocus
   End If
End If
End Sub

Private Sub EliminaCuenta(sCod As String, nItem As Integer)
EliminaRow fgDetalle, fgDetalle.Row
EliminaFgObj nItem
If Len(fgDetalle.TextMatrix(1, 1)) > 0 Then
   RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.Row, 0)
End If
End Sub

Private Sub mnuNoAtender_Click()
Dim nPos As Integer, nSalto As Integer
nSalto = IIf(fgDetalle.Row < fgDetalle.RowSel, 1, -1)
For nPos = fgDetalle.Row To fgDetalle.RowSel Step nSalto
    fgDetalle.TextMatrix(nPos, 5) = ""
Next
SumasDoc
End Sub

Private Sub ActualizaFG(nItem As Integer)
fgDetalle.TextMatrix(nItem, 1) = sObjCod
fgDetalle.TextMatrix(nItem, 2) = sObjDesc
fgDetalle.TextMatrix(nItem, 3) = sObjUnid
fgDetalle.TextMatrix(nItem, 8) = sCtaCod
fgDetalle.TextMatrix(nItem, 9) = sCtaProvis
If fgDetalle.TextMatrix(nItem, 11) = "" Then
   fgDetalle.TextMatrix(nItem, 11) = txtPlazo
End If
fgDetalle.Col = 4
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdDocumento_Click()
Dim oPrevio As clsprevio
Dim lsCadena As String
Dim sPrenombre As String
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim i As Integer
Dim K As Integer
Dim div As Integer
Dim j As Integer
Dim Desc As String
Dim DescF As String
Dim sGlosa As String
Dim sGlosaF As String

Dim nSubt As Currency
Dim nIGVT As Currency
Dim lineas As Long
Dim lsCabecera As String

If chkIGV.value = 1 And chkRenta4.value = 1 Then
   MsgBox "Solo debe elegir 1 concepto IGV / Renta 4ta", vbInformation, "Aviso"
   Exit Sub
End If

Set oPrevio = New clsprevio

lsNegritaOn = oImpresora.gPrnBoldON
lsNegritaOff = oImpresora.gPrnBoldOFF
lsCabecera = ""

If Right(cmbTipoOC.Text, 1) = "D" Then
    sPrenombre = " DIRECTA "
ElseIf Right(cmbTipoOC.Text, 1) = "P" Then
   sPrenombre = " PROCESO "
End If

lsCabecera = lsCabecera & "." & Space(8) & lsNegritaOn & "CMAC MAYNAS S.A." & lsNegritaOff

If lbBienes Then
   If Right(cmbTipoOC.Text, 1) = "D" Or Right(cmbTipoOC.Text, 1) = "P" Then
        lsCabecera = lsCabecera & Space(66) & lsNegritaOn & "ORDEN DE COMPRA " & lsNegritaOff & "- " & sPrenombre & oImpresora.gPrnSaltoLinea
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
  
Dim rs1 As ADODB.Recordset
Dim lnRucProv As String
Set rs1 = New ADODB.Recordset

Set rs1 = GetProveedorRUC(txtPersona.Text)
If Not (rs1.EOF And rs1.BOF) Then
        lnRucProv = rs1!cPersIDnro
End If

  lsCadena = lsCadena & Space(10) & lsNegritaOn & "PROVEEDOR : " & JIZQ(txtProvNom.Text, 45) & lsNegritaOff & Space(3) & lsNegritaOn & "RUC:  " & lsNegritaOff & JIZQ(lnRucProv, 12)
  lsCadena = lsCadena & Space(4) & "Iquitos, " & JIZQ(ArmaFecha(gdFecSis), 40) & oImpresora.gPrnSaltoLinea
  lsCadena = lsCadena & Space(10) & lsNegritaOn & "DIRECCION : " & lsNegritaOff & JIZQ(txtProvDir.Text, 45) & Space(3) & lsNegritaOn & "Telef.: " & lsNegritaOff & JIZQ(txtProvTele.Text, 12) & oImpresora.gPrnSaltoLinea

  lsCadena = lsCadena & Space(10) & lsNegritaOn & "Sirvase atender de acuerdo al siguiente detalle" & oImpresora.gPrnSaltoLinea

  lsCadena = lsCadena & Space(10) & String(115, "_") & oImpresora.gPrnSaltoLinea
  lsCadena = lsCadena & Space(10) & Space(1) & "Cantidad" & Space(1) & "|" & Space(2) & "Unidad" & Space(2) & "|" & Space(1) & "Descripcion" & Space(50) & "|" & Space(1) & "P.Vta.Unit" & " |" & Space(3) & "P.Vta.Total" & Space(1) & "|" & oImpresora.gPrnSaltoLinea
  lsCadena = lsCadena & Space(10) & String(115, "_") & lsNegritaOff & oImpresora.gPrnSaltoLinea

  lineas = 8
  
  i = 1

  Do While Not fgDetalle.Rows - 1
    If fgDetalle.TextMatrix(i, 1) = "" Then
       Exit Do
    End If
    Dim lsCantidadCad As String
    Dim lsTotalCad As String
                
    lsCantidadCad = fgDetalle.TextMatrix(i, 5)
    lsTotalCad = fgDetalle.TextMatrix(i, 7)
  
     Desc = JustificaTextoCadenaOrdenCompra(fgDetalle.TextMatrix(i, 2), 60, 33)


           Dim a As Integer
           a = TextoFinLen
           
           If lbBienes Then
              lsCadena = lsCadena & Space(13) & fgDetalle.TextMatrix(i, 4) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(i, 4)))) & Space(1) & fgDetalle.TextMatrix(i, 3) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(i, 3)))) & Space(3) & Trim(Desc) & Space(60 - a) & _
                          JDER(Trim(fgDetalle.TextMatrix(i, 5)), 14) & JDER(Trim(fgDetalle.TextMatrix(i, 7)), 15) & oImpresora.gPrnSaltoLinea

            Else
               If Len(Desc) > 60 Then
                    lsCadena = lsCadena & Space(13) & fgDetalle.TextMatrix(i, 4) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(i, 4)))) & Space(1) & fgDetalle.TextMatrix(i, 3) & Space(8 - Len(RTrim(fgDetalle.TextMatrix(i, 3)))) & Space(2) & Trim(Desc) & Space(62 - a) & _
                                JDER(Trim(fgDetalle.TextMatrix(i, 5)), 14) & JDER(Trim(fgDetalle.TextMatrix(i, 7)), 15) & oImpresora.gPrnSaltoLinea
               Else
                    lsCadena = lsCadena & Space(13) & Space(19) & JIZQ(Desc, 60) & JDER(fgDetalle.TextMatrix(i, 7), 32) & oImpresora.gPrnSaltoLinea
               End If
            End If
       
    i = i + 1
    lineas = lineas + 1
  Loop
  '*****************************************
lineas = lineas + i

  lsCadena = lsCadena & Space(10) & String(115, "_") & oImpresora.gPrnSaltoLinea
  
  Dim lsMonto As String * 74
  lsMonto = Trim(CStr(ConvNumLet(nVal(txtTot.Text))))
  Dim nSubRenta As Currency
  Dim nRenta4 As Currency
  
  If lbBienes Then 'Orden Compra
    nSubt = Format(((txtTot.Text) / 1.19), "0.00")
    nIGVT = txtTot.Text - nSubt

        If chkIGV.value = 1 Then
           lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(Format(nSubt, gcFormView), 18) & oImpresora.gPrnSaltoLinea
           lsCadena = lsCadena & Space(15) & Space(74) & "I.G.V." & JDER(Format(nIGVT, gcFormView), 27) & oImpresora.gPrnSaltoLinea
           lineas = lineas + 2
        Else
           lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(Format(txtTot.Text, gcFormView), 18) & oImpresora.gPrnSaltoLinea
           lineas = lineas + 1
        End If
     
        lsCadena = lsCadena & Space(15) & Space(97) & String(11, "_") & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & Space(10) & "SON: " & lsMonto & "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(txtTot.Text, 18) & oImpresora.gPrnSaltoLinea
        lineas = lineas + 2
  Else
    If chkRenta4.value = 1 Then
       nSubRenta = Format(((CCur(txtTot.Text)) / 10), "0.00")
       nRenta4 = txtTot.Text - nSubRenta
       
       lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(Format(txtTot.Text, gcFormView), 20) & oImpresora.gPrnSaltoLinea
       lsCadena = lsCadena & Space(15) & Space(74) & "Ret. 4ta." & JDER(Format(nSubRenta, gcFormView), 26) & oImpresora.gPrnSaltoLinea
       lsCadena = lsCadena & Space(15) & Space(98) & String(11, "_") & oImpresora.gPrnSaltoLinea
       lsCadena = lsCadena & Space(10) & "SON: " & lsMonto & "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(Format(nRenta4, gcFormView), 20) & oImpresora.gPrnSaltoLinea
       lineas = lineas + 4
     Else
         If chkIGV.value = 1 Then
            nSubt = Format(((txtTot.Text) / 1.19), "0.00")
            nIGVT = txtTot.Text - nSubt
         
            lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(Format(nSubt, gcFormView), 20) & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Space(15) & Space(74) & "I.G.V." & JDER(Format(nIGVT, gcFormView), 29) & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Space(15) & Space(98) & String(11, "_") & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Space(10) & "SON: " & lsMonto & "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(txtTot.Text, 20) & oImpresora.gPrnSaltoLinea
            lineas = lineas + 4
         Else
            lsCadena = lsCadena & Space(15) & Space(74) & "V.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(Format(txtTot.Text, gcFormView), 20) & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Space(15) & Space(98) & String(11, "_") & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & Space(10) & "SON: " & lsMonto & "P.Vta.Total " & IIf(Mid(gcOpeCod, 3, 1) = 1, "S/.", "$. ") & JDER(txtTot.Text, 20) & oImpresora.gPrnSaltoLinea
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
       
     For j = 0 To div
         If j = 0 Then
            sGlosaF = Mid(sGlosa, 1, 110) & oImpresora.gPrnSaltoLinea & Space(10)
         End If
         If j = div Then
            sGlosaF = sGlosaF & Mid(sGlosa, (110 * j) + 1, 90)
         End If
         If j > 0 And j < div Then
            sGlosaF = sGlosaF & Mid(sGlosa, (110 * j) + 1, 110) & oImpresora.gPrnSaltoLinea & Space(38)
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
          lsCadena = lsCadena & Space(10) & "   técnicas mínimas requeridas, y de acuerdo a lo orfertado por el Contratista, así como de anular la Orden de" & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "   Compra, según el procedimiento establecido de acuerdo a Ley" & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "3. El Contratista se compromete a cumplir las obligaciones que le corresponden, bajo sanción de quedar" & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "   inhabilitado para contratar con el Estado en caso de incumplimiento." & oImpresora.gPrnSaltoLinea
          lsCadena = lsCadena & Space(10) & "4. En caso de retraso injustificado en la entrega de los bienes, se aplicarán las penalidades de acuerdo a Ley." & oImpresora.gPrnSaltoLinea
          lineas = lineas + 10
       End If
       
       For K = lineas To 40
           lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
       Next
    
       lsCadena = lsCadena & Space(10) & String(115, "_") & oImpresora.gPrnSaltoLinea
       lsCadena = lsCadena & Space(14) & "LOGISTICA" & Space(35) & "GERENCIA" & Space(45) & "PROVEEDOR" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    
       lsCadena = lsCadena & Space(10) & String(115, "_") & oImpresora.gPrnSaltoLinea
       lsCadena = lsCadena & Space(10) & Space(10) & "Cualquier consulta: Teléfono (065) 22-3323 / 22-1256  Anexos 1132 - 1195 - 1229 / Fax Anexo 1196" & oImpresora.gPrnSaltoLinea
       lsCadena = lsCadena & Space(10) & Space(15) & "Correos: labreu@cmacmaynas.com.pe,  cpanduro@cmacmaynas.com.pe" & oImpresora.gPrnSaltoLinea
       lsCadena = lsCadena & Space(10) & String(115, "_") & oImpresora.gPrnSaltoLinea

  

oPrevio.Show lsCadena, Caption, , , gIBM
End Sub

Private Sub fgDetalle_DblClick()
If lbImprime Then
   Exit Sub
End If
If fgDetalle.Col = 1 Then
   EnfocaTexto txtObj, 0, fgDetalle
End If
If (fgDetalle.Col = 4 And fgDetalle.TextMatrix(fgDetalle.Row, 10) = "B") Or (fgDetalle.Col = 5 And fgDetalle.TextMatrix(fgDetalle.Row, 10) = "B") Or fgDetalle.Col = 7 Then
   EnfocaTexto txtCant, 0, fgDetalle
End If
If fgDetalle.Col = 2 Then
   EnfocaTexto txtConcepto, 0, fgDetalle
End If
If fgDetalle.Col = 11 Then
   EnfocaTexto txtFecPlazo, 0, fgDetalle
End If
End Sub

Private Sub Form_Activate()
    Dim ofun As NContFunciones
    Set ofun = New NContFunciones
    GetTipCambioLog gdFecSis, Not gbBitCentral
    
    If Not Trim(txtProy.Tag) = "" Then
       chkProy.value = vbChecked
    End If
    If lSalir Then
       Unload Me
    End If
    
    'If Me.cmbMoneda.Text = "" Then
        If Mid(gcOpeCod, 3, 1) = gcMNDig Then
            cmbMoneda.ListIndex = 1
        Else
            cmbMoneda.ListIndex = 0
        End If
    'Else
    '    cmbMoneda.ListIndex = 1
    'End If
    If Me.txtProvCod.Text = "" And Not lbModifica Then
        gcDocNro = ""
    End If
    lTransActiva = False
       
    
    If gcDocNro = "" And gcDocTpo <> "" Then
   
        If Me.cmbMoneda.Text <> "" Then gcDocNro = ofun.GeneraDocNro(CInt(gcDocTpo), Right(Me.cmbMoneda.Text, 1), Year(gdFecSis))
        lblDoc.Caption = UCase(sDocDesc) & "   Nº " & gcDocNro
        
    End If
    
End Sub

Private Sub Form_Load()
    Dim N As Integer, nSaldo As Currency, K As Currency
    Dim sCtaCod As String, nItem As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oConst As COMDConstantes.DCOMConstantes
    Set oConst = New COMDConstantes.DCOMConstantes
    Dim oArea As DActualizaDatosArea
    Set oArea = New DActualizaDatosArea
    Dim oPer As UPersona
    Set oPer = New UPersona
    Dim sTipodoc  As String
    
    Dim ofun As NContFunciones
    Set ofun = New NContFunciones
    
   
    lSalir = False
    CargaComboLog oConst.GetConstante(gMoneda), Me.cmbMoneda
    
    lOrdenCompra = True
    
    If gcOpeCod = "501207" Or gcOpeCod = "502207" Then
        cmdSeleccion.Visible = False
    End If
    
    If Mid(gcOpeCod, 3, 1) = gcMEDig Then  'Identificación de Tipo de Moneda
       gsSimbolo = gcME
       If gnTipCambio = 0 Then
            GetTipCambioLog gdFecSis, Not gbBitCentral
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
       gsSimbolo = gcMN
    End If
    
    Dim ObjOperacion As COMNAuditoria.DOperacion
    Set ObjOperacion = New COMNAuditoria.DOperacion
    Set rs = ObjOperacion.CargaOpeCta(gcOpeCod, "H")
    
    If rs.EOF Then
       MsgBox "No se asignó Cuenta de Provisión de Bienes a Operación", vbInformation, "¡Aviso!"
       rs.Close
       Set rs = Nothing
       lSalir = True
       Exit Sub
    End If
    sCtaProvis = rs!cCtaContCod
    
    txtFecha = Format(gdFecSis, gsFormatoFechaView)
    
  
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
          lblDoc.BackColor = &H80FF80
          txtTot.BackColor = &H80FF80
       End If
       
    Else
       txtObj.BackColor = lnColorServ
       txtCant.BackColor = lnColorServ
       fgDetalle.BackColor = lnColorServ
       fgDetalle.BackColorBkg = lnColorServ
       lbltipooc.Caption = "Tipo O/S"
       
       If Mid(gcOpeCod, 3, 1) = 2 Then
          lblDoc.BackColor = &H80FF80
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
         
         sSql = " SELECT  mc.nMovNro, mc.nMovItem, mo.nMovObjOrden cMovObjOrden, mc.cCtaContCod, " _
              & " mo.cObjetoCod, mcd.cDescrip, dItemPlazo, moc.nMovCant," & IIf(gsSimbolo = gcMN, "mc.nMovImporte ", "me.nMovMeImporte ") & " nMovImporte " _
              & " FROM " & sBaseNew & "MovCta mc " & IIf(gsSimbolo = gcMN, "", " JOIN " & sBaseNew & "MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem") _
              & " LEFT JOIN ( SELECT nMovNro,nMovItem, nMovObjOrden, cObjetoCod FROM " & sBaseNew & "MovObj " _
              & "             WHERE nMovNro = '" & gcMovNro & "') mo on mo.nMovNro = mc.nMovNro and mo.nMovItem = mc.nMovItem" _
              & " LEFT JOIN " & sBaseNew & "MovCant moc ON moc.nMovNro = mo.nMovNro and moc.nMovItem = mo.nMovItem" _
              & " LEFT JOIN " & sBaseNew & "MovCotizacDet mcd ON mcd.nMovNro = mc.nMovNro and mcd.nMovItem = mc.nMovItem " _
              & " WHERE mc.nMovNro = '" & gcMovNro & "' and mc.nMovImporte <> 0 And mc.cCtaContCod Not Like  '25%' ORDER BY mc.nMovItem "
         
         Set rs = oCon.CargaRecordSet(sSql)
         N = 0
         nItem = 0
         Do While Not rs.EOF
            If nItem <> rs!nMovItem Then
               N = N + 1
               If N > fgDetalle.Rows - 1 Then
                  AdicionaRow fgDetalle
               End If
               fgDetalle.TextMatrix(N, 0) = N
               fgDetalle.TextMatrix(N, 1) = rs!cCtaContCod
               fgDetalle.TextMatrix(N, 7) = Format(rs!nMovImporte, gsFormatoNumeroView)
               fgDetalle.TextMatrix(N, 8) = rs!cCtaContCod
               If Not IsNull(rs!dItemPlazo) Then
                  fgDetalle.TextMatrix(N, 11) = rs!dItemPlazo
               End If
               If Not IsNull(rs!cDescrip) Then
                  fgDetalle.TextMatrix(N, 2) = rs!cDescrip
               End If
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
                  If fgDetalle.TextMatrix(N, 2) = "" Then
                     fgDetalle.TextMatrix(N, 2) = rsObj!cObjetoDesc
                  End If
                  If Not IsNull(rsObj!nMovCant) Then
                     fgDetalle.TextMatrix(N, 1) = rsObj!cObjetoCod
                     fgDetalle.TextMatrix(N, 4) = rsObj!nMovCant
                     fgDetalle.TextMatrix(N, 3) = rsObj!unidad
                     If rs!nMovCant <> 0 Then
                        fgDetalle.TextMatrix(N, 5) = Format(Round(rs!nMovImporte / rsObj!nMovCant, 2), gcFormView)
                     End If
                     fgDetalle.TextMatrix(N, 10) = "B"
                     
                     FlexBackColor fgDetalle, N, lnColorBien
                     rs.MoveNext
                  Else
                     nItem = rs!nMovItem
                     fgDetalle.TextMatrix(N, 1) = rs!cCtaContCod
                     fgDetalle.TextMatrix(N, 10) = "S"
                     FlexBackColor fgDetalle, N, lnColorServ
                     Do While Not rsObj.EOF
                        If Not IsNull(rsObj!cObjetoCod) Then
                           AdicionaRow fgObj
                           fgObj.TextMatrix(fgObj.Row, 0) = N
                           fgObj.TextMatrix(fgObj.Row, 1) = rs!cMovObjOrden
                           fgObj.TextMatrix(fgObj.Row, 2) = rsObj!cObjetoCod
                           fgObj.TextMatrix(fgObj.Row, 3) = rsObj!cObjetoDesc
                           fgObj.TextMatrix(fgObj.Row, 5) = sObjCod
                        End If
                        rsObj.MoveNext
                     Loop
                     rs.MoveNext
                  End If
               Else
                  rs.MoveNext
               End If
            Else
               fgDetalle.TextMatrix(N, 10) = "S"
               FlexBackColor fgDetalle, N, lnColorServ
               rs.MoveNext
            End If
         Loop
       
       SumasDoc
    Else
        Dim Obj As COMNAuditoria.DOperacion
        Set Obj = New COMNAuditoria.DOperacion
       Set rs = Obj.CargaOpeDocEstado(gcOpeCod, "1", "_2")
       If RSVacio(rs) Then
          MsgBox "No se definió Documento ORDEN DE COMPRA en Operación. Por favor Consultar con Sistemas...!", vbInformation, "¡Aviso!"
          lSalir = True
          Exit Sub
       Else
          gcDocTpo = rs!nDocTpo
          If Me.cmbMoneda.Text <> "" Then gcDocNro = ofun.GeneraDocNro(CInt(gcDocTpo), Right(Me.cmbMoneda.Text, 1), Year(gdFecSis))
       End If
       For N = 1 To fgDetalle.Rows - 1
          If lbBienes Then
             fgDetalle.TextMatrix(N, 10) = "B"
          Else
             fgDetalle.TextMatrix(N, 10) = "S"
          End If
       Next
       fgDetalle.Row = 1
       fgDetalle.Col = 1
       fgDetalle.Row = 1: fgDetalle.Col = 1
       txtPlazo.value = gdFecSis
    End If
    
    'If Me.cmbMoneda.Text <> "" Then gcDocNro = GeneraDocNro(CInt(gcDocTpo), Right(Me.cmbMoneda.Text, 1), Year(gdFecSis))
    
    lblDoc.Caption = UCase(sDocDesc) & "   Nº " & gcDocNro
    
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
    If lbImprime Then
       cmdDocumento.Visible = True
       cmdAceptar.Visible = False
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
    Me.txtCodLugarEntrega.rs = oConst.getAgencias(, , True)
    txtPersona.TipoBusPers = BusPersDocumentoRuc
    
    
End Sub

Private Sub TxtArea_EmiteDatos()
    Me.txtAgeDesc.Text = txtArea.psDescripcion
    If txtArea.psDescripcion <> "" Then
        Me.txtCodLugarEntrega.SetFocus
    End If
End Sub


Private Sub txtConcepto_GotFocus()
fEnfoque txtConcepto
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   fgDetalle.TextMatrix(txtConcepto.Tag, 2) = txtConcepto.Text
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
      fgDetalle.TextMatrix(txtFecPlazo.Tag, 11) = txtFecPlazo
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

Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   txtMovDesc.SetFocus
End If
End Sub

Private Sub txtPlazo_LostFocus()
Dim K As Integer
For K = 1 To fgDetalle.Rows - 1
   If fgDetalle.TextMatrix(K, 11) = "" And fgDetalle.TextMatrix(K, 1) <> "" Then
      fgDetalle.TextMatrix(K, 11) = txtPlazo
   Else
      If fgDetalle.TextMatrix(K, 1) = "" Then
         fgDetalle.TextMatrix(K, 11) = ""
      End If
   End If
Next
End Sub

Private Sub txtObj_GotFocus()
cmdExaminar.Visible = True
cmdExaminar.Top = txtObj.Top + 15
cmdExaminar.Left = txtObj.Left + txtObj.Width - cmdExaminar.Width
End Sub


Private Sub txtCant_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCant, KeyAscii, 12, 2)
If KeyAscii = 13 Then
   fgDetalle.Text = Format(txtCant.Text, gcFormView)
   If fgDetalle.Col = 4 Then
      fgDetalle.TextMatrix(fgDetalle.Row, 7) = Format(Round(Val(txtCant) * Val(Format(fgDetalle.TextMatrix(fgDetalle.Row, 5), "#0.00")), 2), gcFormView)
   End If
   If fgDetalle.Col = 5 Then
      fgDetalle.TextMatrix(fgDetalle.Row, 7) = Format(Round(Val(txtCant) * Val(Format(fgDetalle.TextMatrix(fgDetalle.Row, 4), "#0.00")), 2), gcFormView)
   End If
   If fgDetalle.Col = 7 And Val(fgDetalle.TextMatrix(fgDetalle.Row, 4)) <> 0 Then
      fgDetalle.TextMatrix(fgDetalle.Row, 5) = Format(Round(Val(txtCant) / Val(Format(fgDetalle.TextMatrix(fgDetalle.Row, 4), "#0.00")), 2), gcFormView)
   End If
   txtCant.Visible = False
   fgDetalle.SetFocus
   SumasDoc
End If
End Sub
Private Sub SumasDoc()
Dim N As Integer
Dim nTot As Currency
For N = 1 To fgDetalle.Rows - 1
    If fgDetalle.TextMatrix(N, 7) <> "" Then
       nTot = nTot + Val(Format(fgDetalle.TextMatrix(N, 7), gcFormDato))
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

Public Property Get lOk() As Boolean
    lOk = OK
End Property

Public Property Let lOk(ByVal vNewValue As Boolean)
OK = vNewValue
End Property

Private Function ValidaAgencia(sAgeCod As String) As Boolean
End Function

Private Function ValidaDatos() As Boolean
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
    If txtLugarEntrega = "" And lbBienes Then
       MsgBox "Falta indicar el Lugar de Entrega de la Orden", vbInformation, "¡Aviso!"
       txtCodLugarEntrega.SetFocus
       Exit Function
    End If
    If Me.cmbMoneda.Text = "" Then
       MsgBox "Falta indicar la moneda", vbInformation, "¡Aviso!"
       cmbMoneda.SetFocus
       Exit Function
    End If
    ValidaDatos = True
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
fgDetalle.TextMatrix(0, 1) = "Objeto"
fgDetalle.TextMatrix(0, 2) = "Descripción"
fgDetalle.TextMatrix(0, 3) = "Unidad"
fgDetalle.TextMatrix(0, 4) = "Solicitado"
fgDetalle.TextMatrix(0, 5) = "P.Unitario"
fgDetalle.TextMatrix(0, 6) = "Saldo"
fgDetalle.TextMatrix(0, 7) = "Sub Total"
fgDetalle.TextMatrix(0, 11) = "Plazo"
fgDetalle.ColWidth(0) = 335
fgDetalle.ColWidth(1) = 1500
fgDetalle.ColWidth(2) = 3500
fgDetalle.ColWidth(3) = 880
fgDetalle.ColWidth(4) = 1200
fgDetalle.ColWidth(5) = 1200
fgDetalle.ColWidth(6) = 0
fgDetalle.ColWidth(7) = 1200
fgDetalle.ColWidth(8) = 0
fgDetalle.ColWidth(9) = 0
fgDetalle.ColWidth(10) = 0
fgDetalle.ColWidth(11) = 1200

fgDetalle.ColAlignment(1) = 1
fgDetalle.ColAlignmentFixed(0) = 4
fgDetalle.ColAlignmentFixed(4) = 7
fgDetalle.ColAlignmentFixed(5) = 7
fgDetalle.ColAlignmentFixed(6) = 7
fgDetalle.ColAlignmentFixed(7) = 7
fgDetalle.ColAlignmentFixed(11) = 4
fgDetalle.RowHeight(-1) = 285
End Sub

Private Sub EliminaFgObj(nItem As Integer)
Dim K  As Integer, m As Integer
K = 1
Do While K < fgObj.Rows
   If Len(fgObj.TextMatrix(K, 1)) > 0 Then
      If Val(fgObj.TextMatrix(K, 0)) = nItem Then
         EliminaRow fgObj, K
         K = K - 1
      Else
         K = K + 1
      End If
   Else
      K = K + 1
   End If
Loop
End Sub

Private Sub AdicionaObj(sCodCta As String, nFila As Integer, rs As ADODB.Recordset)
   Dim nItem As Integer
   AdicionaRow fgObj
   nItem = fgObj.Row
   fgObj.TextMatrix(nItem, 0) = nFila
   fgObj.TextMatrix(nItem, 1) = rs!nCtaObjOrden
End Sub

Private Sub EliminaObjeto(nItem As Integer)
    EliminaFgObj nItem
    If Len(fgDetalle.TextMatrix(1, 1)) > 0 Then
       RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.Row, 0)
    End If
End Sub

Private Sub RefrescaFgObj(nItem As Integer)
    Dim K  As Integer
    For K = 1 To fgObj.Rows - 1
        If Len(fgObj.TextMatrix(K, 1)) Then
           If fgObj.TextMatrix(K, 0) = nItem Then
              fgObj.RowHeight(K) = 285
           Else
              fgObj.RowHeight(K) = 0
           End If
        End If
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

