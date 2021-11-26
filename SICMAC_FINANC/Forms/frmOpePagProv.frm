VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5F774E03-DB36-4DFC-AAC4-D35DC9379F2F}#1.1#0"; "VertMenu.ocx"
Begin VB.Form frmOpePagProv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Operaciones: Pago de Comprobantes Pendientes"
   ClientHeight    =   6465
   ClientLeft      =   645
   ClientTop       =   1635
   ClientWidth     =   11415
   Icon            =   "frmOpePagProv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkAfectoITF 
      Caption         =   "Afecto a ITF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8340
      TabIndex        =   42
      Top             =   4995
      Value           =   1  'Checked
      Width           =   1845
   End
   Begin VB.CheckBox chkNC 
      Appearance      =   0  'Flat
      Caption         =   "NC"
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
      Height          =   285
      Left            =   4095
      TabIndex        =   41
      Top             =   5910
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCtaBcos 
      Caption         =   "Cta Bcos."
      Height          =   345
      Left            =   6945
      TabIndex        =   28
      Top             =   5955
      Width           =   1095
   End
   Begin VB.OptionButton optTipo 
      Appearance      =   0  'Flat
      Caption         =   "Pago a Proveedores"
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
      Height          =   195
      Index           =   1
      Left            =   3495
      TabIndex        =   1
      Top             =   -30
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.OptionButton optTipo 
      Appearance      =   0  'Flat
      Caption         =   "Pago a Policias"
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
      Height          =   195
      Index           =   0
      Left            =   1155
      TabIndex        =   0
      Top             =   -30
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame FraFechaMov 
      Height          =   495
      Left            =   9375
      TabIndex        =   25
      Top             =   -75
      Width           =   1980
      Begin MSMask.MaskEdBox txtFechaMov 
         Height          =   300
         Left            =   870
         TabIndex        =   26
         Top             =   150
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   -2147483635
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mov. Al..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   90
         TabIndex        =   27
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10020
      TabIndex        =   3
      Top             =   585
      Width           =   1065
   End
   Begin VB.TextBox txtTipoCambio 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   6630
      TabIndex        =   20
      Top             =   5490
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   420
      Left            =   1140
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   5310
      Width           =   10245
   End
   Begin MSComctlLib.ImageList imgRec 
      Left            =   8265
      Top             =   5865
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpePagProv.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpePagProv.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpePagProv.frx":14BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpePagProv.frx":17D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpePagProv.frx":1AF2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   10305
      TabIndex        =   10
      Top             =   5955
      Width           =   1095
   End
   Begin VertMenu.VerticalMenu vFormPago 
      Height          =   6225
      Left            =   120
      TabIndex        =   12
      Top             =   90
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   10980
      MenuCaption1    =   "Forma Pago"
      MenuItemsMax1   =   6
      MenuItemIcon11  =   "frmOpePagProv.frx":1E0C
      MenuItemCaption11=   "Efectivo"
      MenuItemIcon12  =   "frmOpePagProv.frx":2126
      MenuItemCaption12=   "Carta"
      MenuItemIcon13  =   "frmOpePagProv.frx":2440
      MenuItemCaption13=   "Orden Pago"
      MenuItemIcon14  =   "frmOpePagProv.frx":275A
      MenuItemCaption14=   "Cheque"
      MenuItemIcon15  =   "frmOpePagProv.frx":2A74
      MenuItemCaption15=   "Abono"
      MenuItemIcon16  =   "frmOpePagProv.frx":2D8E
      MenuItemCaption16=   "Penalidad"
   End
   Begin RichTextLib.RichTextBox rtxtAsiento 
      Height          =   315
      Left            =   10500
      TabIndex        =   14
      Top             =   6180
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmOpePagProv.frx":30A8
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
   Begin VB.CommandButton cmdDoc 
      Caption         =   "&Emitir"
      Height          =   345
      Left            =   9180
      TabIndex        =   8
      Top             =   5955
      Width           =   1095
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "&Rechazar"
      Height          =   345
      Left            =   9180
      TabIndex        =   9
      Top             =   5715
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   8070
      TabIndex        =   6
      Top             =   5955
      Width           =   1095
   End
   Begin Sicmact.FlexEdit fg 
      Height          =   3765
      Left            =   1140
      TabIndex        =   16
      Top             =   1170
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6641
      Cols0           =   27
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmOpePagProv.frx":3129
      EncabezadosAnchos=   "400-0-500-2100-1140-4000-0-1250-0-0-0-0-0-0-1400-1500-2000-2300-1500-1800-1200-2500-2500-0-0-0-0"
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
      ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-C-L-L-R-L-C-C-C-C-C-L-R-R-R-R-L-R-R-R-R-C-C-R"
      FormatosEdit    =   "0-0-0-0-0-0-0-2-0-0-0-0-0-0-0-2-0-2-2-0-2-2-2-4-4-4-3"
      TextArray0      =   "Nro."
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame fraEntidad 
      Caption         =   "Entidad Pagadora"
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
      Left            =   1140
      TabIndex        =   13
      Top             =   5700
      Visible         =   0   'False
      Width           =   5775
      Begin Sicmact.TxtBuscar txtBuscaEntidad 
         Height          =   345
         Left            =   150
         TabIndex        =   15
         Top             =   210
         Width           =   1875
         _ExtentX        =   3307
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
      Begin VB.Label lblCtaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2010
         TabIndex        =   7
         Top             =   210
         Width           =   3705
      End
   End
   Begin VB.CommandButton cmdGenerarArch 
      Caption         =   "&Archivo"
      Height          =   345
      Left            =   8070
      TabIndex        =   17
      Top             =   5715
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkTodos 
      Appearance      =   0  'Flat
      Caption         =   "Todos"
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
      Height          =   285
      Left            =   1185
      TabIndex        =   18
      Top             =   915
      Visible         =   0   'False
      Width           =   975
   End
   Begin VertMenu.VerticalMenu vFormPagoAge 
      Height          =   6225
      Left            =   60
      TabIndex        =   29
      Top             =   90
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   10980
      MenuCaption1    =   "Forma Pago"
      MenuItemIcon11  =   "frmOpePagProv.frx":3248
      MenuItemCaption11=   "Orden Pago"
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Documentos Emitidos"
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
      Left            =   1155
      TabIndex        =   11
      Top             =   195
      Visible         =   0   'False
      Width           =   5175
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   315
         Left            =   810
         TabIndex        =   37
         Top             =   270
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   315
         Left            =   3000
         TabIndex        =   38
         Top             =   270
         Width           =   1155
         _ExtentX        =   2037
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
         PromptChar      =   "_"
      End
      Begin VB.Frame FraTipoB 
         Height          =   885
         Left            =   2700
         TabIndex        =   30
         Top             =   1005
         Width           =   5295
         Begin VB.OptionButton optAge 
            Appearance      =   0  'Flat
            Caption         =   "Todos"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   165
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optAge 
            Appearance      =   0  'Flat
            Caption         =   "x Agencia"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   32
            Top             =   165
            Width           =   1065
         End
         Begin VB.OptionButton optAge 
            Appearance      =   0  'Flat
            Caption         =   "Logistica"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   31
            Top             =   165
            Width           =   1125
         End
         Begin Sicmact.TxtBuscar txtAge 
            Height          =   345
            Left            =   930
            TabIndex        =   34
            Top             =   450
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   609
            Appearance      =   0
            BackColor       =   14811132
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
            EnabledText     =   0   'False
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Agencia:"
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
            Left            =   150
            TabIndex        =   36
            Top             =   510
            Width           =   765
         End
         Begin VB.Label lblAgencia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1950
            TabIndex        =   35
            Top             =   450
            Width           =   3255
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2340
         TabIndex        =   40
         Top             =   330
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   195
         TabIndex        =   39
         Top             =   285
         Width           =   555
      End
   End
   Begin VB.Frame fraBuscar 
      Height          =   705
      Left            =   1140
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   10215
      Begin Sicmact.TxtBuscar txtProveedor 
         Height          =   375
         Left            =   3540
         TabIndex        =   4
         Top             =   195
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   661
         Appearance      =   0
         BackColor       =   14811132
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
         EnabledText     =   0   'False
      End
      Begin VB.Frame fraFiltro 
         Height          =   495
         Left            =   90
         TabIndex        =   23
         Top             =   105
         Width           =   2190
         Begin MSMask.MaskEdBox mskFiltro 
            Height          =   300
            Left            =   1155
            TabIndex        =   2
            Top             =   150
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   -2147483635
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Filtro:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   105
            TabIndex        =   24
            Top             =   180
            Width           =   990
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Documento Enviado:"
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
         Height          =   435
         Left            =   2520
         TabIndex        =   22
         Top             =   165
         Width           =   1065
      End
   End
   Begin VB.Label lblTC 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Cambio:"
      Height          =   195
      Left            =   5430
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   1155
   End
End
Attribute VB_Name = "frmOpePagProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCtaContDebeB As String
Dim lsCtaContDebeS As String
Dim lsCtaContDebeRH As String
Dim lsCtaContDebeRHJ As String
Dim lsCtaContDebeSegu As String
Dim lsCtaContDebePagoVarios As String
Dim lsCtaContDebeBLeasingMN As String 'ALPA 20110911
Dim lsCtaContDebeBLeasingME As String 'ALPA 20110911
Dim lsCtaContDebeBLeasingSMN As String 'ALPA 20110911
Dim lsCtaContDebeBLeasingSME As String 'ALPA 20110911

Dim lsDocs      As String
Dim lsFileCarta As String
Dim lsDocTpo    As String

Dim rs As New ADODB.Recordset
Dim lSalir As Boolean
Dim lmn As Boolean
Dim lTransActiva As Boolean
Dim ContSalOp   As Integer
Dim sFileOP     As String
Dim sFileVE     As String
Dim sFileNA     As String
Dim sFileVNA    As String
Dim fs          As Scripting.FileSystemObject
Dim oBarra      As clsProgressBar
Dim lsPersId As String
Dim lbReporte As Boolean
Dim lsCaption As String
Dim lsTipoB As String
Dim lsCtaITFD As String
Dim lsCtaITFH As String
'Dim objPista As COMManejador.Pista
Dim oOpe        As DOperacion
'ALPA 20090317*****************************************************************
Dim lnTipoPago As Integer
Dim lnTipoDocTemp As Integer
Dim lsCtaRetencion As String

'ALPA 20110909
Dim lsCtaCore As String
Dim lsCtaSAF As String


'******************************************************************************

Public Sub Ini(pbReporte As Boolean, psCaption As String)
    lbReporte = pbReporte
    lsCaption = psCaption
    Me.Show 1
End Sub

Private Function ValidaInterfaz() As Boolean
ValidaInterfaz = False
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Falta indicar Descripción de Operación", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Exit Function
End If
If fraEntidad.Visible Then
    If txtBuscaEntidad.Text = "" Then
        MsgBox "Cuenta de Institución Financiera no válida", vbInformation, "Aviso"
        Exit Function
    End If
End If
ValidaInterfaz = True
End Function

Private Function BuscaCtaProveedor(psCodPers As String, psMoneda As String) As String
    Dim oCon As New DConecta ' ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim sSql As String
    
    sSql = " select ccodcta from CGPagoProveedorCta where cperscod='" & psCodPers & "' and cmoneda='" & psMoneda & "'"
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    If Not RSVacio(rs) Then
        BuscaCtaProveedor = rs!cCodCta
    End If
End Function

Private Sub chkTodos_Click()
    Dim i As Integer
    If Me.chkTodos.value = 1 Then
        For i = 1 To Me.fg.Rows - 1
            'If Me.fg.TextMatrix(I, 16) <> "RUC NO REGISTRADO" Then
                Me.fg.TextMatrix(i, 2) = 1
            'End If
        Next i
    Else
        For i = 1 To Me.fg.Rows - 1
            'Call fg_OnCellCheck(i, 1)
            Me.fg.TextMatrix(i, 2) = 0
        Next i
    End If
End Sub

Private Sub cmdCtaBcos_Click()
'    If Me.fg.TextMatrix(fg.Row, 8) = "" Then Exit Sub
'    frmACGCtasProveedores.Ini Me.fg.TextMatrix(fg.Row, 8), Me.fg.TextMatrix(fg.Row, 5)
'    Set frmACGCtasProveedores = Nothing
End Sub

Private Sub cmdDoc_Click()
Dim K                As Integer
Dim lsEntidadOrig    As String
Dim lsCtaEntidadOrig As String
Dim lsPersNombre  As String
Dim lsPersDireccion As String
Dim lsUbigeo    As String
Dim lsCuentaAho As String
Dim lbGrabaOpeNegocio As Boolean
Dim lnImporteB    As Currency
Dim lnImporteS    As Currency
Dim lnImporteAjusteDolaresS As Currency
Dim lnImporteAjusteDolaresB As Currency
Dim lnImporteAjusteDolaresRH As Currency
Dim lnImporteAjusteDolaresRHJ As Currency
Dim lnImporteAjusteDolaresSEGU As Currency
Dim lnImporteAjusteDolaresPagoVarios As Currency

Dim oDocPago      As clsDocPago
Dim lsSubCuentaIF As String
Dim lsPersCod     As String
Dim lsDocNRo      As String
Dim lsMovNro      As String
Dim lsDocumento   As String
Dim lsOpeCod      As String
Dim lsCtaBanco    As String
Dim lsCtaContHaber As String
Dim lsPersCodIF    As String
Dim lsMovAnt       As String
Dim lnMovAnt       As Long
Dim lsCtaContHaberGen As String
Dim rsBilletaje As ADODB.Recordset
Dim lsOPSave    As String
Dim lbEfectivo  As Boolean
Dim oNContFunc  As NContFunciones
Dim oDCtaIF     As DCajaCtasIF
Dim oCtasIF     As NCajaCtaIF
Dim lsTpoIf     As String
Dim oOpe        As DOperacion
Dim lsDocVoucher As String
Dim lsFecha      As String
'Dim lnITFValor As Currency
Dim lnITFValor As Double '*** PEAC 20110331

Dim lnMontoDif  As Currency
Dim lsCtaDiferencia As String
Dim nDocs       As Integer
Dim lsCabeImpre As String
Dim lsImpre     As String
Dim lsCadBol    As String

Dim oNCaja As nCajaGeneral
Set oNCaja = New nCajaGeneral

'Retencion
Dim oConst As NConstSistemas
Set oConst = New NConstSistemas
Dim oImpuesto As DImpuesto
Set oImpuesto = New DImpuesto

Dim lbBitReten As Boolean
Dim lbDocConIGV As Boolean
Dim lsCtaReten As String
Dim lbBCAR As Boolean
Dim lnTasaImp As Currency
Dim lnIngresos As Currency
Dim lnRetencion As Currency
Dim lnTopeRetencion As Currency
Dim lnRetAct As Currency
Dim lnRetActME As Currency
Dim lsComprobante As String
Dim oPrevio As clsPrevioFinan
Set oPrevio = New clsPrevioFinan
Dim lnDocProv As String
Dim lsCtaContDebeRH As String
Dim lsCtaContDebeRHJ As String
'Dim lsCtaContDebeSegu As String
Dim lsCtaContDebePagoVarios As String
Dim lnImporteRH    As Currency
Dim lnImporteRHJ    As Currency
Dim lnImporteSEGU    As Currency
Dim lnImportePAGOVARIOS    As Currency
Dim lbOk        As Boolean
Dim lsCtaContLeasing As String
'Dim lnITF As Currency
Dim lnITF As Double '*** PEAC 20110331
Dim nLogico As Integer

Dim lnMontoPago As Currency



lbBitReten = IIf(oConst.LeeConstSistema(gConstSistBitRetencion6Porcent) = 1, True, False)

On Error GoTo NoGrabo
If txtFechaMov.Enabled Then
    If Not ValidaFechaContab(txtFechaMov, gdFecSis) Then
        Exit Sub
    End If
End If
If ValidaInterfaz = False Then Exit Sub
If lsDocTpo = "" Then
    MsgBox "Seleccione Forma de Pago", vbInformation, "Aviso"
    Exit Sub
End If

If lsDocTpo = "" Then
    MsgBox "Seleccione Forma de Pago", vbInformation, "Aviso"
    Exit Sub
End If
If lsDocTpo = 48 And chkAfectoITF.value = 0 Then ' Orden de pago
    MsgBox "Orden de Pago debe ser Afecto a ITF", vbInformation, "Aviso"
    Exit Sub
End If
If lsDocTpo = 58 And chkAfectoITF.value = 0 Then 'Abono en cuenta
    MsgBox "Abono en Cuenta debe ser Afecto a ITF", vbInformation, "Aviso"
    Exit Sub
End If



Set oNContFunc = New NContFunciones
Set oCtasIF = New NCajaCtaIF
Set oDCtaIF = New DCajaCtasIF
Set oOpe = New DOperacion
Set oDocPago = New clsDocPago

lsCtaEntidadOrig = Trim(lblCtaDesc)
lsTpoIf = Mid(txtBuscaEntidad, 1, 2)
lsCtaBanco = Mid(txtBuscaEntidad, 18, Len(Me.txtBuscaEntidad))
lsPersCodIF = Mid(txtBuscaEntidad, 4, 13)
lsEntidadOrig = oDCtaIF.NombreIF(lsPersCodIF)
lsSubCuentaIF = oCtasIF.SubCuentaIF(lsPersCodIF)

gsGlosa = Trim(txtMovDesc)
'lsMovAnt = fg.TextMatrix(fg.Row, 9)
'lnMovAnt = fg.TextMatrix(fg.Row, 10)

lsDocVoucher = ""
lsDocNRo = ""
lbEfectivo = False
lsCadBol = ""
lsCtaContHaber = ""
lsPersCod = ""
lnImporteB = 0: lnImporteS = 0
lnImporteAjusteDolaresS = 0
lnImporteAjusteDolaresB = 0
lsCabeImpre = " DOCUMENTOS PAGADOS : "
nDocs = 0


'For K = 1 To fg.Rows - 1
'   If fg.TextMatrix(K, 2) = "." Then
'      nDocs = nDocs + 1
'      If Not lsPersCod = "" Then
'         If Not lsPersCod = fg.TextMatrix(fg.Row, 8) Then
'            MsgBox "No se puede hacer Pago a Proveedores Diferentes", vbInformation, "¡Aviso!"
''            fg.SetFocus
'            Exit Sub
'         End If
'      End If
'      lsPersCod = fg.TextMatrix(fg.Row, 8)
'      lsPersNombre = fg.TextMatrix(fg.Row, 5)
'      lsCabeImpre = lsCabeImpre & oImpresora.gPrnCondensadaON & fg.TextMatrix(K, 3) & Space(5) & oImpresora.gPrnCondensadaOFF
'      If nDocs Mod 4 = 0 Then
'         lsCabeImpre = lsCabeImpre & oImpresora.gPrnSaltoLinea & Space(22)
'      End If
'
''      If Me.optTipo(0) = True Then
''        If lsCtaContDebeB = fg.TextMatrix(K, 13) Then
''           lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 15))
''           lnImporteAjusteDolaresB = lnImporteAjusteDolaresB '+ CCur(fg.TextMatrix(k, 20))
''        End If
''        If lsCtaContDebeS = fg.TextMatrix(K, 13) Then
''           lnImporteS = lnImporteS + CCur(fg.TextMatrix(K, 15))
''           lnImporteAjusteDolaresS = lnImporteAjusteDolaresS '+ CCur(fg.TextMatrix(k, 20))
''        End If
''      Else
''        If lsCtaContDebeB = fg.TextMatrix(K, 13) Then
''           lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 18))
''           lnImporteAjusteDolaresB = lnImporteAjusteDolaresB + CCur(fg.TextMatrix(K, 20))
''        End If
''        If lsCtaContDebeS = fg.TextMatrix(K, 13) Then
''           lnImporteS = lnImporteS + CCur(fg.TextMatrix(K, 18))
''           lnImporteAjusteDolaresS = lnImporteAjusteDolaresS + CCur(fg.TextMatrix(K, 20))
''        End If
''      End If
'
'
'
'
'   End If
'Next

For K = 1 To fg.Rows - 1
   If fg.TextMatrix(K, 2) = "." Then
   'ALPA 20090318*******************************************************************
      If fg.TextMatrix(K, 26) <> lnTipoPago Then
            MsgBox "No se puede realizar diferentes tipos de pagos", vbInformation, "¡Aviso!"
            fg.SetFocus
            Exit Sub
      End If
    '********************************************************************************
      nDocs = nDocs + 1
      If Not lsPersCod = "" Then
         If Not lsPersCod = fg.TextMatrix(fg.Row, 8) Then
            MsgBox "No se puede hacer Pago a Proveedores Diferentes", vbInformation, "¡Aviso!"
            fg.SetFocus
            Exit Sub
         End If
      End If
      lsPersCod = fg.TextMatrix(fg.Row, 8)
      lsPersNombre = fg.TextMatrix(fg.Row, 5)
      
'      If Me.optTipo(0) = True Then
'        If lsCtaContDebeB = fg.TextMatrix(K, 13) Then
'           lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 15))
'           lnImporteAjusteDolaresB = lnImporteAjusteDolaresB '+ CCur(fg.TextMatrix(k, 20))
'        End If
'        If lsCtaContDebeS = fg.TextMatrix(K, 13) Then
'           lnImporteS = lnImporteS + CCur(fg.TextMatrix(K, 15))
'           lnImporteAjusteDolaresS = lnImporteAjusteDolaresS '+ CCur(fg.TextMatrix(k, 20))
'        End If
'      Else
'        If lsCtaContDebeB = fg.TextMatrix(K, 13) Then
'           lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 18))
'           lnImporteAjusteDolaresB = lnImporteAjusteDolaresB + CCur(fg.TextMatrix(K, 20))
'        End If
'        If lsCtaContDebeS = fg.TextMatrix(K, 13) Then
'           lnImporteS = lnImporteS + CCur(fg.TextMatrix(K, 18))
'           lnImporteAjusteDolaresS = lnImporteAjusteDolaresS + CCur(fg.TextMatrix(K, 20))
'        End If
'      End If
      
      If gsCodCMAC = 106 Then
         lsCabeImpre = lsCabeImpre & fg.TextMatrix(K, 3) & space(5)
         lnDocProv = lnDocProv & fg.TextMatrix(K, 3) & space(5)
      Else
         lsCabeImpre = lsCabeImpre & oImpresora.gPrnCondensadaON & fg.TextMatrix(K, 3) & space(5) & oImpresora.gPrnCondensadaOFF
      End If
      If nDocs Mod 4 = 0 Then
         lsCabeImpre = lsCabeImpre & oImpresora.gPrnSaltoLinea & space(22)
      End If
      nLogico = 0
      If lsCtaContDebeB = fg.TextMatrix(K, 13) And Not (lsCtaContDebeBLeasingMN = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSMN = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingME = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSME = fg.TextMatrix(K, 13)) Then
         lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 18))
         lnImporteAjusteDolaresB = lnImporteAjusteDolaresB + CCur(fg.TextMatrix(K, 20))
      End If
      If lsCtaContDebeS = fg.TextMatrix(K, 13) Then
         lnImporteS = lnImporteS + CCur(fg.TextMatrix(K, 18))
         lnImporteAjusteDolaresS = lnImporteAjusteDolaresS + CCur(fg.TextMatrix(K, 20))
      End If
      
      If lsCtaContDebeRH = fg.TextMatrix(K, 13) Then
         lnImporteRH = lnImporteRH + CCur(fg.TextMatrix(K, 18))
         lnImporteAjusteDolaresRH = lnImporteAjusteDolaresRH + CCur(fg.TextMatrix(K, 20))
      End If
      If lsCtaContDebeRHJ = fg.TextMatrix(K, 13) Then
         lnImporteRHJ = lnImporteRHJ + CCur(fg.TextMatrix(K, 18))
         lnImporteAjusteDolaresRHJ = lnImporteAjusteDolaresRHJ + CCur(fg.TextMatrix(K, 20))
      End If
      If lsCtaContDebeSegu = fg.TextMatrix(K, 13) Then
         lnImporteSEGU = lnImporteSEGU + CCur(fg.TextMatrix(K, 18))
         lnImporteAjusteDolaresSEGU = lnImporteAjusteDolaresSEGU + CCur(fg.TextMatrix(K, 20))
      End If
      
      If lsCtaContDebePagoVarios = fg.TextMatrix(K, 13) Then
         lnImportePAGOVARIOS = lnImportePAGOVARIOS + CCur(fg.TextMatrix(K, 18))
         lnImporteAjusteDolaresPagoVarios = lnImporteAjusteDolaresPagoVarios + CCur(fg.TextMatrix(K, 20))
      End If
      
      'ALPA 20110911 -------------------------------------
       If lsCtaContDebeBLeasingMN = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSMN = fg.TextMatrix(K, 13) Then
         lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 18))
         lnImporteAjusteDolaresB = lnImporteAjusteDolaresB + CCur(fg.TextMatrix(K, 20))
         lsCtaContLeasing = fg.TextMatrix(K, 13)
      End If
      
      If lsCtaContDebeBLeasingME = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSME = fg.TextMatrix(K, 13) Then
         lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 18))
         lnImporteAjusteDolaresB = lnImporteAjusteDolaresB + CCur(fg.TextMatrix(K, 20))
         lsCtaContLeasing = fg.TextMatrix(K, 13)
      End If
      '---------------------------------------------------
      
      If lbBitReten And Not lbDocConIGV Then
         'Segun Jose es solo para Facturas
         If Val(fg.TextMatrix(K, 11)) = TpoDocFactura Or Val(fg.TextMatrix(K, 11)) = TpoDocNotaCredito Or Val(fg.TextMatrix(K, 11)) = TpoDocNotaDebito Then
            lbDocConIGV = True
         End If
      End If
   End If
Next


If lnImporteB + lnImporteS + lnImporteSEGU = 0 Then
   MsgBox "No se Seleccionó Comprobantes para Pagar!", vbInformation, "¡Aviso!"
   fg.SetFocus
   Exit Sub
End If
   
'If gbPermisoLogProveedorAG = True Then
'   Dim oGen As DGeneral
'   Set oGen = New DGeneral
'   Dim lnLogPoderenRIni As Double
'   Dim lnLogPoderenRFin As Double
'
'   lnLogPoderenRIni = oGen.GetParametro(5000, 1006) * oGen.GetParametro(4000, 1010) / 100
'   lnLogPoderenRFin = oGen.GetParametro(5000, 1007) * oGen.GetParametro(4000, 1010) / 100
'   If Not ((lnImporteB + lnImporteS) > lnLogPoderenRIni And CCur((lnImporteB + lnImporteS)) <= lnLogPoderenRFin) Then
'      MsgBox "Monto a pagar por Fondo Fijo no se encuentra en el rango autorizado ", vbInformation, "Aviso"
'      Exit Sub
'   End If
'End If

If lbBitReten Then
    lbBCAR = VerifBCAR(lsPersCod)
    lsCtaReten = oConst.LeeConstSistema(gConstSistCtaRetencion6Porcent)
    lnTasaImp = oImpuesto.CargaImpuesto(lsCtaReten)!nImpTasa
    lnIngresos = oNCaja.GetMontoIngresoRetencion(lsPersCod, Left(Format(gdFecSis, gsFormatoMovFecha), 6), True)
    lnRetencion = oNCaja.GetMontoIngresoRetencion(lsPersCod, Left(Format(gdFecSis, gsFormatoMovFecha), 6), False)
    lnTopeRetencion = oConst.LeeConstSistema(gConstSistTopeRetencion6Porcent)
    
    If Not lbBCAR Then
        If lmn Then
            lnRetAct = (lnImporteB + lnImporteS) + lnIngresos
        Else
            lnRetAct = Round((lnImporteB + lnImporteS) * gnTipCambioPonderado, 2) + lnIngresos
        End If
        If lnRetAct <= lnTopeRetencion Then
           lnRetAct = 0
        Else
           lnRetAct = Round(lnRetAct * (lnTasaImp / 100), 2) - lnRetencion
           
           If lmn Then
              If lnRetAct > (lnImporteB + lnImporteS) Then
                 lnRetAct = (lnImporteB + lnImporteS)
              End If
           Else
              If lnRetAct > (lnImporteB + lnImporteS) * gnTipCambioPonderado Then
                 lnRetAct = (lnImporteB + lnImporteS) * gnTipCambioPonderado
              End If
           End If
        End If
    Else
        lnRetAct = 0
    End If

    If lmn Then
        lnRetActME = 0
    Else
        lnRetActME = Round(lnRetAct / gnTipCambioPonderado, 2)
    End If
    
   If lnRetAct > 0 Then  'Proveedor esta Afecto a Retencion
      Dim sTexto As String
      Dim N      As Integer
      Do While True
         sTexto = InputBox("El proveeedor: " & lsPersNombre & " esta afecto a una retención de : ", "Retención a Pago", Round(lnRetAct, 2))
         If sTexto = "" Then
            Exit Do
         End If
         If IsNumeric(sTexto) Then
            lnRetAct = CCur(sTexto)
            Exit Do
         Else
            MsgBox "Debe ingresar dato Númerico", vbInformation, "¡Aviso!"
         End If
      Loop
      lnRetActME = Round(lnRetAct / gnTipCambioPonderado, 2)
   End If
   If lnRetAct > 0 Then
       If MsgBox("El proveeedor: " & lsPersNombre & " esta afecto a una retención de (" & gcMN & ") : " & Format(lnRetAct, "#,##0.00") & vbNewLine & "Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
           Exit Sub
       End If
   End If
Else
    lnRetAct = 0
End If
'And gbPermisoLogProveedorAG = False
If lsDocTpo = "-1" And lnTipoPago = 0 Then
    frmArendirEfectivo.Inicio 0, fg.TextMatrix(fg.Row, 12), Mid(gsOpeCod, 3, 1), "", lnImporteB + lnImporteS + lnImporteSEGU - IIf(lmn, lnRetAct, lnRetActME), lsPersCod, lsPersNombre, ArendirRendicion, "Nro.Doc.:"
    If frmArendirEfectivo.vnDiferencia <> 0 Then
        lnMontoDif = frmArendirEfectivo.vnDiferencia
    End If
    Set rsBilletaje = frmArendirEfectivo.rsEfectivo
    Set frmArendirEfectivo = Nothing
    If rsBilletaje Is Nothing Then
        Exit Sub
    End If
    lbEfectivo = True
    lsFecha = Format(gdFecSis, gsFormatoFechaView)
ElseIf lsDocTpo = TpoDocNotaAbono Then
    Dim oImp As New NContImprimir
    Dim lsCtaAho As String
    lsDocTpo = TpoDocNotaAbono
'    lsCtaAho = BuscaCtaProveedor(lsPersCod, Mid(gsOpeCod, 3, 1))
    frmNotaCargoAbono.Inicio lsDocTpo, lnImporteB + lnImporteS + lnImporteSEGU - IIf(lmn, lnRetAct, lnRetActME), gdFecSis, txtMovDesc, gsOpeCod, False, lsPersCod, lsPersNombre, lsCtaAho, , lnITFValor
    If frmNotaCargoAbono.vbOk Then
        lsDocNRo = frmNotaCargoAbono.NroNotaCA
        txtMovDesc = frmNotaCargoAbono.Glosa
        lsDocumento = frmNotaCargoAbono.NotaCargoAbono
        lsPersNombre = frmNotaCargoAbono.PersNombre
        lsPersDireccion = frmNotaCargoAbono.PersDireccion
        lsUbigeo = frmNotaCargoAbono.PersUbigeo
        lsCuentaAho = frmNotaCargoAbono.CuentaAhoNro
        lsFecha = frmNotaCargoAbono.FechaNotaCA
        lnITF = frmNotaCargoAbono.lnITFValor

'        lsDocumento = oImp.ImprimeNotaCargoAbono(lsDocNRo, txtMovDesc, CCur(frmNotaCargoAbono.Monto), _
'                            lsPersNombre, lsPersDireccion, lsUbigeo, gdFecSis, Mid(gsOpeCod, 3, 1), lsCuentaAho, lsDocTpo, gsNomAge, gsCodUser)
         lsDocumento = oImp.ImprimeNotaAbono(lsFecha, lnImporteB + lnImporteS + lnImporteSEGU - IIf(lmn, lnRetAct, lnRetActME), txtMovDesc, lsCuentaAho, lsPersNombre)
        lbGrabaOpeNegocio = MsgBox(" ¿ Desea que se realice Abono en Cuenta del Proveedor ? ", vbQuestion + vbYesNo, "¡Confirmacion!") = vbYes
        If lbGrabaOpeNegocio Then
            Dim oDis As New NRHProcesosCierre
            lsCadBol = oDis.ImprimeBoletaCad(CDate(lsFecha), "ABONO CAJA GENERAL", "Depósito CAJA GENERAL*Nro." & lsDocNRo, "", lnImporteB + lnImporteS + lnImporteSEGU - IIf(lmn, lnRetAct, lnRetActME), lsPersNombre, lsCuentaAho, "", 0, 0, "Nota Abono", 0, 0, False, False, , , , True, , , , False, gsNomAge) & oImpresora.gPrnSaltoPagina
        End If
    Else
        Exit Sub
    End If
Else
    If lsDocTpo = TpoDocCheque Then
       lsDocVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1))
       'oDocPago.InicioCheque lsDocNRo, True, Mid(Me.txtBuscaEntidad, 4, 13), gsOpeCod, lsPersNombre, gsOpeDesc, gsGlosa, lnImporteB + lnImporteS + lnImporteSEGU - IIf(lMN, lnRetAct, lnRetActME), gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, True, gsCodAge
       oDocPago.InicioCheque lsDocNRo, True, Mid(Me.txtBuscaEntidad, 4, 13), gsOpeCod, lsPersNombre, gsOpeDesc, gsGlosa, lnImporteB + lnImporteS + lnImporteSEGU - IIf(lmn, lnRetAct, lnRetActME), gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, True, gsCodAge, , , lsTpoIf, lsPersCodIF, lsCtaBanco 'EJVG20121129
    End If
    If lsDocTpo = TpoDocOrdenPago Then
       lsDocVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1))
       oDocPago.InicioOrdenPago lsDocNRo, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeCod, gsGlosa, lnImporteB + lnImporteS + lnImporteSEGU - IIf(lmn, lnRetAct, lnRetActME), gdFecSis, lsDocVoucher, True
    End If
    If lsDocTpo = TpoDocCarta Then
       oDocPago.InicioCarta lsDocNRo, lsPersCod, gsOpeCod, gsOpeCod, gsGlosa, lsFileCarta, lnImporteB + lnImporteS + lnImporteSEGU - lnRetAct, gdFecSis, lsEntidadOrig, lsCtaEntidadOrig, lsPersNombre, "", lsMovNro, gnMgDer, gnMgIzq, gnMgSup
    End If
'    'ALPA 20090323******************************************
    If lsDocTpo = TpoDocRetenciones Then
        oDocPago.InicioPenalidad lsDocNRo, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeCod, gsGlosa, lnImporteB + lnImporteS + lnImporteSEGU - IIf(lmn, lnRetAct, lnRetActME), gdFecSis, lsDocVoucher, True
    End If
'    '*******************************************************
    If oDocPago.vbOk Then    'Se ingresó dato de Cheque u Orden de Pago
       lsOpeCod = gsOpeCod
       lsFecha = oDocPago.vdFechaDoc
       lsDocTpo = oDocPago.vsTpoDoc
       lsDocNRo = oDocPago.vsNroDoc
       lsDocVoucher = oDocPago.vsNroVoucher
       lsDocumento = oDocPago.vsFormaDoc
        If lsDocTpo = TpoDocCarta Then
            gnMgIzq = oDocPago.vnMargIzq
            gnMgDer = oDocPago.vnMargDer
            gnMgSup = oDocPago.vnMargSup
        End If
    Else
'        MsgBox " Seleccione el tipo de modalidad de pago ", vbInformation, "Aviso"
       Exit Sub
    End If
End If
'ALPA 20110909*****************************
Dim oCtSaldo As DCtaSaldo
Set oCtSaldo = New DCtaSaldo
Dim oRS As ADODB.Recordset
Set oRS = New ADODB.Recordset
Set oRS = oCtSaldo.ObtenerOperacionesSAF(txtProveedor.Text)
    lsCtaCore = ""
    lsCtaSAF = ""
If Not (oRS.BOF Or oRS.EOF) Then
    lsCtaCore = oRS!cCtaCod
    lsCtaSAF = oRS!cCtaSaf
End If
'******************************************
'ALPA 20090403*****************************
lnTipoDocTemp = lsDocTpo
If lsDocTpo = TpoDocRetenciones Then
   lsDocTpo = -1
End If
If lnTipoDocTemp = TpoDocRetenciones Then

    If Mid(gsOpeCod, 3, 1) = 1 Then
       lsCtaContHaber = "62" & Mid(gsOpeCod, 3, 1) & "10909" & Right(gsCodAge, 2)
    Else
       lsCtaContHaber = "62" & Mid(gsOpeCod, 3, 1) & "10909" & Right(gsCodAge, 2)
    End If
Else
'******************************************
    lsOpeCod = oOpe.EmiteOpeDoc(Mid(gsOpeCod, 1, 5), lsDocTpo)
    If lsOpeCod = "" Then
       MsgBox "No se asignó Documentos de Referencia a Operación de Pago", vbInformation, "Aviso"
       Exit Sub
    End If

    If lsDocTpo = TpoDocOrdenPago Then
    '   If gbPermisoLogProveedorAG = True Then
    '    lsCtaContHaber = oOpe.EmiteOpeCta(lsOpeCod, "H", , gsCodArea + gsCodAge, ObjCMACAgenciaArea)
    '   Else
        lsCtaContHaber = oOpe.EmiteOpeCta(lsOpeCod, "H", , gsCodArea, ObjCMACAgenciaArea)
    '   End If
    Else
       lsCtaContHaber = oOpe.EmiteOpeCta(lsOpeCod, "H", , txtBuscaEntidad, ObjEntidadesFinancieras)
    End If
'ALPA 20090403*****************************
End If
'******************************************
If lsCtaContDebeB = "" Or lsCtaContDebeS = "" Or lsCtaContHaber = "" Then
        MsgBox "Cuentas Contables no determinadas Correctamente." & oImpresora.gPrnSaltoLinea & "consulte con Sistemas", vbInformation, "Aviso"
   Exit Sub
End If
 
If MsgBox("Desea Grabar la Información", vbYesNo + vbQuestion, "Aviso") = vbYes Then
   cmdDoc.Enabled = False
   lsMovNro = oNContFunc.GeneraMovNro(txtFechaMov, Right(gsCodAge, 2), gsCodUser)
   If lsDocTpo = TpoDocLetras Then
      lbOk = oNCaja.GrabaCanjePorLetras(lsMovNro, lsOpeCod, txtMovDesc, rsBilletaje, lsCtaContDebeB, lsCtaContDebeS, lnImporteB, lnImporteS, fg.GetRsNew, lsPersCod)
   Else
      lsCtaDiferencia = oOpe.EmiteOpeCta(lsOpeCod, IIf(lnMontoDif > 0, "H", "D"), "2")
      If lsOpeCod = "421113" Or lsOpeCod = "422113" Then
         lsCabeImpre = space(8) & lsCabeImpre
      End If
'ALPA 200900318*********************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
'      lbOk = oNCaja.GrabaPagoProveedor(lsMovNro, lsOpeCod, txtMovDesc, lsCtaContDebeB, lsCtaContDebeS, _
'                                   lsCtaContHaber, lnImporteB, lnImporteS, lsPersCod, lsTpoIf, lsPersCodIf, lsCtaBanco, _
'                                   rsBilletaje, lsDocTpo, lsDocNRo, Format(CDate(lsFecha), gsFormatoFecha), lsDocVoucher, fg.GetRsNew, lsCuentaAho, lsCtaDiferencia, lnMontoDif, gbBitCentral, lbGrabaOpeNegocio, IIf(Mid(gsOpeCod, 3, 1) = 1, lnRetAct, lnRetActME), lsCtaITFD, lsCtaITFH, gnImpITF, IIf(chkNC.value = 1, True, False), IIf(chkAfectoITF.value = 1, True, False), lnITF, lnImporteRH, lsCtaContDebeRH, lsCtaContDebeRHJ, lnImporteRHJ, lsCtaContDebeSegu, lnImporteSEGU, lsCtaContDebePagoVarios, lnImportePAGOVARIOS) = 0
    lbOk = oNCaja.GrabaPagoProveedor(lsMovNro, lsOpeCod, txtMovDesc, lsCtaContDebeB, lsCtaContDebeS, _
                                   lsCtaContHaber, lnImporteB, lnImporteS, lsPersCod, lsTpoIf, lsPersCodIF, lsCtaBanco, _
                                   rsBilletaje, lsDocTpo, lsDocNRo, Format(CDate(lsFecha), gsFormatoFecha), lsDocVoucher, fg.GetRsNew, lsCuentaAho, lsCtaDiferencia, lnMontoDif, gbBitCentral, lbGrabaOpeNegocio, IIf(Mid(gsOpeCod, 3, 1) = 1, lnRetAct, lnRetActME), lsCtaITFD, lsCtaITFH, gnImpITF, IIf(chkNC.value = 1, True, False), IIf(chkAfectoITF.value = 1, True, False), lnITF, lnImporteRH, lsCtaContDebeRH, lsCtaContDebeRHJ, lnImporteRHJ, lsCtaContDebeSegu, lnImporteSEGU, lsCtaContDebePagoVarios, lnImportePAGOVARIOS, lnTipoPago, lsCtaSAF, lsCtaCore, lsCtaContLeasing) = 0
'***********************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
   
'   If oNCaja.GrabaPagoProveedor(lsMovNro, lsOpeCod, txtMovDesc, lsCtaContDebeB, lsCtaContDebeS, _
                                   lsCtaContHaber, lnImporteB, lnImporteS, lsPersCod, lsTpoIf, lsPersCodIf, lsCtaBanco, _
                                   rsBilletaje, lsDocTpo, lsDocNRo, Format(CDate(lsFecha), gsFormatoFecha), lsDocVoucher, fg.GetRsNew, lsCuentaAho, lsCtaDiferencia, lnMontoDif, gbBitCentral, lbGrabaOpeNegocio, lnRetAct, , , lnImporteAjusteDolaresB, lnImporteAjusteDolaresS, gsCodAge, gsCodArea) = 0 Then
              
      If lbOk Then
         If lsOpeCod = "421113" Or lsOpeCod = "422113" Then
             ImprimeAsientoContableUltimo lsMovNro, lsDocVoucher, lsDocTpo, lsDocumento, lbEfectivo, False, txtMovDesc, lsPersCod, lnImporteB + lnImporteS + lnImporteRH + lnImporteRHJ + lnImporteSEGU + lnImportePAGOVARIOS - IIf(lmn, lnRetAct, lnRetActME), , , , 1, , "17", , lsCabeImpre, lsCadBol, Mid(lsOpeCod, 3, 1), lnDocProv
         Else
             ImprimeAsientoContable lsMovNro, lsDocVoucher, lsDocTpo, lsDocumento, lbEfectivo, False, txtMovDesc, lsPersCod, lnImporteB + lnImporteS - IIf(lmn, lnRetAct, lnRetActME), , , , 1, , "17", , lsCabeImpre, lsCadBol
         End If
      End If
      
      'objPista.InsertarPista lsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Pago a Proveedores"
      
   If lbOk Then
      K = 1
      lnMontoPago = lnImporteB + lnImporteS + lnImporteRH + lnImporteRHJ + lnImporteSEGU
      Do While K < fg.Rows And lnMontoPago <> 0
         If fg.TextMatrix(K, 2) = "." Then
            If lnMontoPago >= CCur(fg.TextMatrix(K, 7)) Then
               lnMontoPago = lnMontoPago - CCur(fg.TextMatrix(K, 7))
               fg.EliminaFila K
            Else
               fg.TextMatrix(K, 7) = Format(CCur(fg.TextMatrix(K, 7)) - lnMontoPago, gsFormatoNumeroView)
               lnMontoPago = 0
            End If
         Else
            K = K + 1
         End If
      Loop
      cmdDoc.Enabled = True
      Set oNCaja = Nothing
      If lbBitReten And (lnRetAct <> 0 Or lnRetActME <> 0) Then
           lsComprobante = GetComprobRetencion(lsMovNro)
           If lsComprobante <> "" Then
                 oPrevio.Show lsComprobante, Caption, False
           End If
      End If
      If fg.TextMatrix(1, 0) = "" Then
          Unload Me
          Exit Sub
      End If
      txtMovDesc = ""
      lsDocTpo = ""
      lsDocNRo = ""
      lsDocVoucher = ""
      lsDocumento = ""
      txtBuscaEntidad = ""
      lblCtaDesc = ""
   End If
   cmdDoc.Enabled = True
   CargaProveedores "P"
        
'      K = 1
'      Do While K < fg.Rows
'         If fg.TextMatrix(K, 2) = "." Then
'            fg.EliminaFila K
'         Else
'            K = K + 1
'         End If
'      Loop
'      cmdDoc.Enabled = True
'      Set oNCaja = Nothing
'      If lbBitReten And (lnRetAct <> 0 Or lnRetActME <> 0) Then
'           lsComprobante = GetComprobRetencion(lsMovNro)
'           If lsComprobante <> "" Then
'                 oPrevio.Show lsComprobante, Caption, False
'           End If
'      End If
'      If fg.TextMatrix(1, 0) = "" Then
'          Unload Me
'          Exit Sub
'      End If
'      txtMovDesc = ""
'      lsDocTpo = ""
'      lsDocNRo = ""
'      lsDocVoucher = ""
'      lsDocumento = ""
'      txtBuscaEntidad = ""
'      lblCtaDesc = ""
'   End If
'   cmdDoc.Enabled = True

End If
End If

Exit Sub

NoGrabo:
  MsgBox TextErr(Err.Description), vbInformation, "Error de Actualización"
  cmdDoc.Enabled = True
End Sub

Private Function Verificar() As Boolean
    Dim i As Integer
    
    Verificar = False
    
    For i = 1 To Me.fg.Rows - 1
        If Me.fg.TextMatrix(i, 2) = "." Then
            Verificar = True
            Exit For
        End If
    Next i
    
    If Verificar = False Then
        MsgBox "No ha seleccionado ningun proveedor para el archivo ", vbInformation, "Aviso"
        Exit Function
    End If
    
    If Not IsNumeric(Me.txtTipoCambio.Text) Then
        MsgBox " Ingrese el tipo de cambio ", vbInformation, "Aviso"
        txtTipoCambio.SetFocus
        Verificar = False
        Exit Function
    End If
     
    Verificar = True
    
End Function

Private Sub cmdGenerarArch_Click()
Dim psArchivoAGrabar    As String
Dim i                   As Integer
Dim sSql                As String
Dim oCon                As New DConecta
Dim oMov                As DMov
Dim lsMovNro            As String
Dim tmpProveedor        As String
Dim tmpnMovNro          As String
Dim tmpMonto            As String
Dim lbTrans             As Boolean
Dim rs                  As ADODB.Recordset
Set rs = New ADODB.Recordset

On Error GoTo mError

    If Verificar = False Then Exit Sub
    
    AplicaTipoCambio
    
    If MsgBox(" ¿ Desea Generar el Archivo de consulta SUNAT ? ", vbExclamation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    Set oMov = New DMov
    lsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    psArchivoAGrabar = App.path & "\spooler\Proveedor" & lsMovNro & ".txt"

    Me.txtTipoCambio.Enabled = False
    oCon.AbreConexion
    oCon.BeginTrans
    lbTrans = True
    Open psArchivoAGrabar For Output As #1
        For i = 1 To Me.fg.Rows - 1
            If Me.fg.TextMatrix(i, 2) = "." Then
                If tmpProveedor = "" Then
                    tmpMonto = CDbl(Me.fg.TextMatrix(i, 15)) - CDbl(Me.fg.TextMatrix(i, 22))
                    tmpProveedor = Me.fg.TextMatrix(i, 16)
                    tmpnMovNro = ""
                ElseIf tmpProveedor = Me.fg.TextMatrix(i, 16) Then
                    tmpMonto = CDbl(tmpMonto) + (CDbl(Me.fg.TextMatrix(i, 15)) - CDbl(Me.fg.TextMatrix(i, 22)))
                    tmpProveedor = Me.fg.TextMatrix(i, 16)
                Else
                    Print #1, Trim(tmpProveedor) & "|" & Trim(Format(tmpMonto, "0.00")) & "|"
                    tmpMonto = CDbl(Me.fg.TextMatrix(i, 15)) - CDbl(Me.fg.TextMatrix(i, 22))
                    tmpProveedor = Me.fg.TextMatrix(i, 16)
                End If
                
                sSql = " Select max(nItem) nItem From MovControlPagoSunat Where nMovNro = " & Trim(Me.fg.TextMatrix(i, 10)) & ""
                Set rs = oCon.CargaRecordSet(sSql)
                
                If IsNull(rs!nItem) Then
                    sSql = "  insert into movcontrolpagosunat(nMovNro,nitem,cMovNro,ntpocambio,bVigente, bValido) "
                    sSql = sSql & " values (" & Trim(Me.fg.TextMatrix(i, 10)) & ",1, '" & lsMovNro & "'," & IIf(fg.TextMatrix(i, 19) = "DOLARES", Me.txtTipoCambio.Text, "0") & ",1,1)"
                Else
'                    sSQL = "  Update movcontrolpagosunat "
'                    sSQL = sSQL & " Set cMovNro = '" & lsMovNro & "', nTpoCambio = " & IIf(fg.TextMatrix(i, 15) <> fg.TextMatrix(i, 7), Me.txtTipoCambio.Text, "0") & " Where nMovNro = " & Trim(Me.fg.TextMatrix(i, 10)) & " and nitem=" & rs!nItem & ""
                    sSql = "  insert into movcontrolpagosunat(nMovNro,nitem,cMovNro,ntpocambio,bVigente, bValido) "
                    sSql = sSql & " values (" & Trim(Me.fg.TextMatrix(i, 10)) & "," & rs!nItem + 1 & ", '" & lsMovNro & "'," & IIf(fg.TextMatrix(i, 19) = "DOLARES", Me.txtTipoCambio.Text, "0") & ",1,1)"

                End If
                oCon.Ejecutar sSql
                rs.Close
            End If
        Next i
    Print #1, Trim(tmpProveedor) & "|" & Trim(Format(tmpMonto, "0.00")) & "|"
    oCon.CommitTrans
    lbTrans = False
    Close #1
    oCon.CierraConexion
    MsgBox " Archivo Generado satisfactoriamente ", vbInformation, "Aviso"
    Me.cmdGenerarArch.Enabled = False
    CargaProveedores
    Me.cmdGenerarArch.Enabled = True
    Me.txtTipoCambio.Enabled = True
    Exit Sub
mError:
    If lbTrans = True Then
        oCon.RollbackTrans
    End If
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdImprimir_Click()
Dim lsOPSave   As String
Dim lsVEOPSave As String

Dim oPlant  As dPlantilla
Dim oNPlant As NPlantilla

Set oPlant = New dPlantilla
Set oNPlant = New NPlantilla


    If gnDocTpo = TpoDocNotaAbono Then
    Else
        lsOPSave = oNPlant.GetPlantillaDoc(IDPlantillaOP)
        If lsOPSave <> "" Then
            EnviaPrevio lsOPSave, "IMPRESION DE ORDENES DE PAGO", gnLinPage
            oPlant.GrabaPlantilla "OPBatch", "Ordenes de Pago para impresiones en Batch", ""
            
            lsVEOPSave = oNPlant.GetPlantillaDoc(IDPlantillaVOP)
            EnviaPrevio lsVEOPSave, "IMPRESION DE VOUCHERS DE EGRESO", gnLinPage
            oPlant.GrabaPlantilla "OPVEBatch", "Voucher de egresos de Ordenes de Pago para impresiones en Batch", ""
        Else
            MsgBox "Archivo de Ordenes de pago no posee información a Imprimir", vbInformation, "Aviso"
        End If
    End If
End Sub



Private Sub cmdProcesar_Click()
'On Error GoTo ErrProcesar
    
'   chkTodos.Visible = True

   If CDate(txtFechaDel) > CDate(txtFechaAl) Then
      MsgBox "Fecha de Inicio no puede ser Mayor que Fecha final", vbInformation, "Aviso"
      Exit Sub
   End If
   
   If Me.optAge(1).value Then
      If Me.txtAge.Text = "" Then
         MsgBox "Debe elegir una agencia.", vbInformation, "Aviso"
         Me.txtAge.SetFocus
         Exit Sub
      End If
   End If
   
   cmdProcesar.Enabled = False

   If lbReporte = True Or gsOpeCod = OpeCGOpeProvRechazo Then
        CargaProveedores
   ElseIf Me.fraBuscar.Visible = True Then
    'ALPA 20090318*****************************************
       'Me.txtProveedor.rs = CargaObjeto(Trim(Me.mskFiltro.Text))
       Me.txtProveedor.rs = CargaObjeto(Trim(Me.mskFiltro.Text), lnTipoPago)
    '******************************************************
   ElseIf Me.fraBuscar.Visible = False Then
        CargaProveedores ("P")
   End If
   cmdProcesar.Enabled = True
   Me.chkTodos.value = 0
'Exit Sub
'ErrProcesar:
'    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub cmdRechazar_Click()
On Error GoTo ErrSave

Dim oFun As New NContFunciones
If Not oFun.PermiteModificarAsiento(fg.TextMatrix(fg.Row, 9), True, CStr(gdFecSis)) Then
    Exit Sub
End If

If Len(fg.TextMatrix(fg.Row, 0)) > 0 Then
    If MsgBox(" ¿ Seguro de Rechazar comprobante de Proveedor ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
       Dim oMov As New DMov
       oMov.EliminaMov fg.TextMatrix(fg.Row, 9)
       Set oMov = Nothing
       fg.EliminaFila fg.Row
       txtMovDesc = ""
    End If
Else
    MsgBox "No existen datos para Rechazar", vbInformation, "¡Aviso!"
End If
Exit Sub
ErrSave:
   MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fg_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    Dim lsPersIdentif As String
    Dim lnI As Long
    
    If Not lbReporte Then
        If fg.TextMatrix(pnRow, 2) = "." Then
            lsPersIdentif = fg.TextMatrix(pnRow, 8)
            
            For lnI = 1 To Me.fg.Rows - 1
                If fg.TextMatrix(lnI, 2) = "." And fg.TextMatrix(lnI, 8) <> lsPersIdentif Then
                   fg.TextMatrix(lnI, 2) = ""
                End If
            Next lnI
            
            fg.Row = pnRow
        End If
    Else
        'If Me.fg.TextMatrix(pnRow, 16) = "RUC NO REGISTRADO" Then Comentado porque no solo se pagan proveedores con RUC GITU
        '    fg.TextMatrix(pnRow, 2) = 0
        'End If
    End If
End Sub

Private Sub fg_RowColChange()
txtMovDesc = fg.TextMatrix(fg.Row, 6)
End Sub

Private Sub Form_Activate()
If lSalir Then
   Unload Me
End If
End Sub

Private Sub Form_Initialize()
    lbReporte = False
End Sub

Private Sub Form_Load()
Dim nTipCambio As Currency
Dim oAge As New DActualizaDatosArea
Dim rs As New ADODB.Recordset
Dim oOpe As New DOperacion
Set oOpe = New DOperacion

On Error GoTo LoadErr
CentraForm Me

FraTipoB.Visible = False

'Set objPista = New COMManejador.Pista

lSalir = False

txtFechaDel = DateAdd("D", -30, gdFecSis)
txtFechaAl = gdFecSis
txtFechaMov = gdFecSis
lsDocTpo = "-1"
ContSalOp = 0

Me.mskFiltro.Text = Format(gdFecSis, gsFormatoMovFecha)
lsTipoB = "T"
gnDocTpo = 0
gsDocNro = ""
gsGlosa = ""

lsCtaITFD = oOpe.EmiteOpeCta(gsOpeCod, "D", 2)
lsCtaITFH = oOpe.EmiteOpeCta(gsOpeCod, "H", 2)


'Set rs = oOpe.CargaOpeCta(gsOpeCod, "D", "0")
'If Not rs.EOF Then
'   lsCtaContDebeB = rs!cCtaContCod
'End If
'
'Set rs = oOpe.CargaOpeCta(gsOpeCod, "D", "1")
'If Not rs.EOF Then
'   lsCtaContDebeS = rs!cCtaContCod
'End If

'Dim oOpe As New DOperacion
lsCtaContDebeB = oOpe.EmiteOpeCta(gsOpeCod, "D", "0") 'Bienes
lsCtaContDebeS = oOpe.EmiteOpeCta(gsOpeCod, "D", "1") 'Servicios
lsCtaContDebeRH = oOpe.EmiteOpeCta(gsOpeCod, "D", "3") 'RRHH Descuento Planillas
lsCtaContDebeRHJ = oOpe.EmiteOpeCta(gsOpeCod, "D", "4") 'RRHH Judicial
lsCtaContDebeSegu = oOpe.EmiteOpeCta(gsOpeCod, "D", "5") 'Seguros
lsCtaContDebePagoVarios = oOpe.EmiteOpeCta(gsOpeCod, "D", "6") 'Pagos Varios
'ALPA 20120503******************************************************************
lsCtaContDebeBLeasingMN = oOpe.EmiteOpeCtaLeasing("421110", 1) '"251603"
lsCtaContDebeBLeasingSMN = oOpe.EmiteOpeCtaLeasing("421110", 2) '"25160203"
lsCtaContDebeBLeasingME = oOpe.EmiteOpeCtaLeasing("422110", 1) '"252603"
lsCtaContDebeBLeasingSME = oOpe.EmiteOpeCtaLeasing("422110", 2) '"25260203"
'*******************************************************************************
Set rs = oOpe.CargaOpeDoc(gsOpeCod, , OpeDocMetDigitado)
lsDocs = RSMuestraLista(rs, 1)
Set oOpe = Nothing
RSClose rs

If lbReporte Then
   fg.ColWidth(17) = 0
   fg.ColWidth(18) = 0
   fg.ColWidth(20) = 0
ElseIf gsOpeCod = OpeCGOpeProvRechazo Then
   fg.ColWidth(16) = 0
   fg.ColWidth(17) = 0
   fg.ColWidth(18) = 0
   fg.ColWidth(19) = 0
   fg.ColWidth(20) = 0
   fg.ColWidth(22) = 0
   fg.ColWidth(23) = 0
   fg.ColWidth(24) = 0
ElseIf Me.optTipo(0) = True Then
    Me.fg.ColWidth(17) = 0
    Me.fg.ColWidth(18) = 0
    Me.fg.ColWidth(16) = 0
    Me.fg.ColWidth(19) = 0
    Me.fg.ColWidth(20) = 0
    Me.fg.ColWidth(21) = 0

End If

Me.Caption = lsCaption
If gsOpeCod = OpeCGOpeProvRechazo Or lbReporte Then
'''    Me.fraBuscar.Top = 30
'''    Me.fraFecha.Top = 30
'''    Me.fg.Top = 750
'''    Me.txtMovDesc.Top = 4610
'''    Me.cmdDoc.Top = 4995
'''    Me.cmdGenerarArch.Top = 4995
'''    Me.CmdImprimir.Top = 4995
'''    Me.cmdProcesar.Top = 4995
'''    Me.cmdRechazar.Top = 4995
'''    Me.cmdSalir.Top = 4995
''''    Me.Height = 6060
'''    Me.FraFechaMov.Top = 180
'''    Me.cmdProcesar.Top = 270
'''    Me.optTipo(0).Visible = False
'''    Me.optTipo(1).Visible = False
'''    cmdCtaBcos.Visible = False
    If gsOpeCod = OpeCGOpeProvRechazo Then
       vFormPago.Visible = False
       fraFecha.Left = fraFecha.Left - 1020
       fraFecha.Width = fraFecha.Width + 1020
       fg.Left = fg.Left - 1020
       fg.Width = fg.Width + 1020
       txtMovDesc.Left = txtMovDesc.Left - 1020
       txtMovDesc.Width = txtMovDesc.Width + 1020
       cmdDoc.Visible = False
       CmdImprimir.Visible = False
       cmdCtaBcos.Visible = False
       cmdRechazar.Visible = True
       Me.fraFecha.Visible = True
       Me.FraFechaMov.Left = 7000
       Me.cmdProcesar.Left = Me.cmdProcesar.Left - 1020
       Me.FraFechaMov.Visible = True
       Me.FraFechaMov.Top = 140
       
       'JEOM
       Me.chkAfectoITF.Visible = False
       'FIN
       
        Me.fraFecha.Visible = True
        Me.fraBuscar.Visible = False
        Me.FraTipoB.Visible = False
        Me.fraFecha.Height = 680
        
        Me.Label3.Left = 2670
        Me.Label3.Top = 330
        Me.txtFechaDel.Left = 3390
        Me.txtFechaDel.Top = 240
        
        Me.Label4.Left = 4620
        Me.Label4.Top = 330
        Me.txtFechaAl.Top = 240
        Me.txtFechaAl.Left = 5190
        
        Me.cmdProcesar.Top = 630
        Me.cmdProcesar.Left = 7720
        Me.FraFechaMov.Top = 530
        Me.FraFechaMov.Left = 9000
        Me.fg.Top = 1200
        Me.txtMovDesc.Top = 4500
        Me.cmdRechazar.Top = 5000
        Me.cmdSalir.Top = 5000
        Me.Height = 6000
       Exit Sub
       
    ElseIf lbReporte Then
'        Me.Caption = lsCaption
'        Me.vFormPago.Visible = False
'        fraFecha.Left = vFormPago.Left
'        Me.fg.Left = vFormPago.Left
'        Me.fg.Top = Me.Frame1.Height + 500
'        Me.txtMovDesc.Left = vFormPago.Left
'        fraEntidad.Visible = False
'        Me.CmdImprimir.Visible = False
'        Me.cmdDoc.Visible = False
'        Me.cmdRechazar.Visible = False
'        Me.Width = 11505 - 1050
'        Me.FraFechaMov.Visible = False
'        Me.cmdGenerarArch.Visible = True
'        Me.cmdSalir.Left = Me.cmdDoc.Left
        chkTodos.Left = txtMovDesc.Left
'        chkTodos.Visible = True
'        Me.lblTC.Visible = True
'        Me.txtTipoCambio.Visible = True
'        Me.fraBuscar.Visible = False
'        Me.cmdProcesar.Left = 10200 - 1050
'        Me.fraFecha.Visible = True
        Me.fraFecha.Top = 90
        Me.chkAfectoITF.Visible = False
        Me.fg.Top = 1350
'        Me.chkTodos.Top = 4710
        Me.txtMovDesc.Top = 5010
        Me.fraEntidad.Top = 5460
        Me.fraFecha.Left = vFormPago.Left
        Me.fg.Left = vFormPago.Left
        Me.txtMovDesc.Left = vFormPago.Left
        Me.chkTodos.Left = vFormPago.Left
        Me.cmdProcesar.Left = Me.cmdProcesar.Left - vFormPago.Width
        Me.cmdSalir.Left = Me.cmdSalir.Left - vFormPago.Width
        Me.Width = Me.Width - vFormPago.Width
        Me.fraFecha.Visible = True
        Me.FraFechaMov.Visible = False
        Me.cmdProcesar.Top = Me.txtFechaAl.Top
        Me.chkTodos.Visible = True
        Me.Height = vFormPago.Height + 300
        Me.cmdGenerarArch.Top = Me.cmdGenerarArch.Top - 200
        Me.cmdGenerarArch.Visible = True
        Me.cmdSalir.Visible = True
        Me.cmdSalir.Top = Me.cmdGenerarArch.Top
        Me.lblTC.Top = Me.cmdGenerarArch.Top
        Me.txtTipoCambio.Top = Me.cmdGenerarArch.Top
        Me.lblTC.Visible = True
        Me.txtTipoCambio.Visible = True
        Me.CmdImprimir.Visible = False
        Me.cmdDoc.Visible = False
        Me.cmdRechazar.Visible = False
        Me.cmdCtaBcos.Visible = False
    End If
Else
'    Me.optTipo(0).Visible = True
'    Me.optTipo(1).Visible = True
    Me.fraFecha.Visible = False
    Me.fraBuscar.Visible = True
    Me.fraBuscar.Left = Me.fg.Left
    Me.fg.Top = 1110
    Me.fg.Height = 3795
    Me.cmdProcesar.Top = 540

    Me.vFormPagoAge.Visible = False
    Me.vFormPago.Visible = True

End If


lmn = IIf(Mid(gsOpeCod, 3, 1) = Moneda.gMonedaExtranjera, False, True)

lsFileCarta = App.path & "\" & gsDirPlantillas & gsOpeCod & ".TXT"
txtBuscaEntidad.psRaiz = "Cuentas de Instituciones Financieras"
Set oOpe = New DOperacion
txtBuscaEntidad.rs = oOpe.GetRsOpeObj(gsOpeCod, "1")  '  oDCtaIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), gTpoIFBanco + gTpoCtaIFCtaCte + gTpoCtaIFCtaAho)
Set oOpe = Nothing

If lbReporte = False Then
    Me.txtProveedor.rs = CargaObjeto
Else
    chkTodos.Visible = True
End If
'ALPA 200900318********************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
lnTipoPago = 0
'**********************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************************
Exit Sub
LoadErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

'ALPA 20090318*************************************************
'Public Function CargaObjeto(Optional psFiltro As String = "") As ADODB.Recordset
Public Function CargaObjeto(Optional psFiltro As String = "", Optional pnTipoPago As Integer = 0) As ADODB.Recordset
'**************************************************************
   On Error GoTo CargaObjetoErr
   Dim oCon As New DConecta
   Dim psSql As String
   
   If oCon.AbreConexion Then
      'psSql = " select distinct cmovnro,'PROVEEDOR'+ CMOVNRO,len(cmovnro) from movcontrolpagosunat where cmovnro like '" & psFiltro & "%' And bVigente = 1"
      
      ' psSql = "  Select distinct ms.cmovnro,'PROVEEDOR'+ ms.CMOVNRO,len(ms.cmovnro)"
      ' psSql = psSql & "  from movcontrolpagosunat ms  join mov m1 on ms.nmovnro=m1.nmovnro and m1.copecod like '__" & Mid(gsOpeCod, 3, 1) & "%'"
      ' psSql = psSql & "   left join movref mr on mr.nmovnroref=ms.nmovnro"
      ' psSql = psSql & "   left join mov m on mr.nmovnro=m.nmovnro and m.copecod not in ('401581','402581')"
      ' psSql = psSql & "   where ms.cmovnro like '" & psFiltro & "%' And bVigente = 1"
      ' psSql = psSql & "   and mr.nmovnro is null and cmovnrores is not null"
      
      psSql = "  Select distinct ms.cmovnro,'PROVEEDOR'+ ms.CMOVNRO,len(ms.cmovnro)"
      psSql = psSql & "  from movcontrolpagosunat ms  join mov m1 on ms.nmovnro=m1.nmovnro and m1.copecod like '__" & Mid(gsOpeCod, 3, 1) & "%'"
      psSql = psSql & "  left join (SELECT mr.nMovNro, mr.nMovNroRef FROM MovRef mr JOIN Mov m1 ON m1.nMovNro = mr.nMovNro "
      psSql = psSql & "                   WHERE m1.nMovEstado = " & gMovEstContabMovContable & " and m1.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagExtornado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') and RTRIM(ISNULL(mr.cAgeCodRef,'')) = '' "
      'ALPA 20090318****************************************************************
      'psSql = psSql & "                  and m1.cOpeCod not in ('" & gOpeCGOpeBancosRetCtasBancosDetracMN & "','" & gOpeCGOpeBancosRetCtasBancosDetracME & "','" & gOpeCGOpeBancosRetCtasBancosEmbargoMN & "','" & gOpeCGOpeBancosRetCtasBancosEmbargoME & "','" & OpeCGOpeProvPagoPagUNAT & "')) mr ON  mr.nMovNroRef = ms.nMovNro "
      psSql = psSql & "                  and m1.cOpeCod not in ('" & gOpeCGOpeBancosRetCtasBancosDetracMN & "','" & gOpeCGOpeBancosRetCtasBancosDetracME & "','" & gOpeCGOpeBancosRetCtasBancosEmbargoMN & "','" & gOpeCGOpeBancosRetCtasBancosEmbargoME & "','" & OpeCGOpeProvPagoPagUNAT & "') and isnull(mr.nTipoPago,0)=" & pnTipoPago & ") mr ON  mr.nMovNroRef = ms.nMovNro "
      '*****************************************************************************
      psSql = psSql & "   left join mov m on mr.nmovnro=m.nmovnro and m.copecod not in ('401581','402581')"
      psSql = psSql & "   where ms.cmovnro like '" & psFiltro & "%' And bVigente = 1"
      psSql = psSql & "   and mr.nmovnro is null and cmovnrores is not null"
      psSql = psSql & "   and m1.nMovEstado = 10 and m1.nMovFlag =0  "
      'ALPA 20090318****************************************************************
      psSql = psSql & "   order by ms.cmovnro "
      '*****************************************************************************
      Set CargaObjeto = oCon.CargaRecordSet(psSql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaObjetoErr:
    MsgBox Err.Description
End Function

Private Sub CargaProveedores(Optional psTipo As String = "")
Dim rs As New ADODB.Recordset
Dim nItem As Long
Dim nTieneDetra As New DMov
Dim bTieneDetra As Boolean
Dim cCtaDetraTemp As String
Dim nCantTempo As Integer
Dim lsTipoInterfaz As String
Dim lnTCFijo As Currency
Dim lsCtaContDebeBME As String
Dim lsCtaContDebeSME As String
Dim lsCtaContDebeRHME As String
Dim lsCtaContDebeRHJME As String
Dim lsCtaContDebeSeguME As String
Dim lsCtaContDebePagoVariosME As String
Dim lsProveedorLeasingMN As String
Dim lsProveedorLeasingME As String
Dim lsProveedorLeasingSMN As String
Dim lsProveedorLeasingSME As String
Dim oOpe As New DOperacion
Set oOpe = New DOperacion
'ALPA20160418***********************
Dim lsProveedorFEPCMACMN As String
Dim lsProveedorFEPCMACME As String
'***********************************
Dim lsProveedorCtasApertSBS As String 'NAGL INC1712260008 20171227

On Error GoTo ErrCargaProveedores

Dim oTC As New nTipoCambio

lnTCFijo = oTC.EmiteTipoCambio(gdFecSis, TCFijoMes)

fg.Clear
fg.Rows = 2
fg.FormaCabecera

If lbReporte Then
    fg.ColWidth(17) = 0
    fg.ColWidth(18) = 0
    fg.ColWidth(20) = 0
ElseIf gsOpeCod = OpeCGOpeProvRechazo Then
   fg.ColWidth(16) = 0
   fg.ColWidth(17) = 0
   fg.ColWidth(18) = 0
   fg.ColWidth(19) = 0
   fg.ColWidth(20) = 0
   fg.ColWidth(22) = 0
ElseIf Me.optTipo(0) = True Then
    Me.fg.ColWidth(17) = 0
    Me.fg.ColWidth(18) = 0
    Me.fg.ColWidth(16) = 0
    Me.fg.ColWidth(19) = 0
    Me.fg.ColWidth(20) = 0
    Me.fg.ColWidth(21) = 0
    Me.fg.ColWidth(22) = 0
End If


Dim oDCaja As New DCajaGeneral


cCtaDetraTemp = Mid(cCtaDetraccionProvision, 1, 2) & Mid(gsOpeCod, 3, 1) & Mid(cCtaDetraccionProvision, 4, Len(cCtaDetraccionProvision) - 2)
'ALPA 20110912
Dim oOpeL As New DOperacion
Set oOpeL = New DOperacion

lsProveedorLeasingMN = oOpeL.EmiteOpeCtaLeasing("421110", 1) '"251601"
lsProveedorLeasingME = oOpeL.EmiteOpeCtaLeasing("421110", 2) '"25160203"
lsProveedorLeasingSMN = oOpeL.EmiteOpeCtaLeasing("422110", 1) '"25160203"
lsProveedorLeasingSME = oOpeL.EmiteOpeCtaLeasing("422110", 2) '"25260203"

lsProveedorFEPCMACMN = "251702" 'ALPA20160418
lsProveedorFEPCMACME = "252702" 'ALPA20160418

Set oOpeL = Nothing

If gsOpeCod = OpeCGOpeProvPago Or gsOpeCod = OpeCGOpeProvPagoME Then
    lsTipoInterfaz = "PAGO"
ElseIf lbReporte = True Then
    lsTipoInterfaz = "lbReporte"
    Me.cmdGenerarArch.Enabled = False
    lsCtaContDebeBME = Left(lsCtaContDebeB, 2) & "2" & Right(lsCtaContDebeB, Len(lsCtaContDebeB) - 3)
    lsCtaContDebeSME = Left(lsCtaContDebeS, 2) & "2" & Right(lsCtaContDebeS, Len(lsCtaContDebeS) - 3)
    lsCtaContDebeRHME = Left(lsCtaContDebeRH, 2) & "2" & Right(lsCtaContDebeRH, Len(lsCtaContDebeRH) - 3) 'RRHH Descuento Planillas
    lsCtaContDebeRHJME = Left(lsCtaContDebeRHJ, 2) & "2" & Right(lsCtaContDebeRHJ, Len(lsCtaContDebeRHJ) - 3) 'RRHH Judicial
    lsCtaContDebeSeguME = Left(lsCtaContDebeSegu, 2) & "2" & Right(lsCtaContDebeSegu, Len(lsCtaContDebeSegu) - 3) 'Seguros
    lsCtaContDebePagoVariosME = Left(lsCtaContDebePagoVarios, 2) & "2" & Right(lsCtaContDebePagoVarios, Len(lsCtaContDebePagoVarios) - 3) 'Pagos Varios
ElseIf gsOpeCod = OpeCGOpeProvRechazo Then
    lsTipoInterfaz = "RECHAZO"
    If Mid(gsOpeCod, 3, 1) = "2" Then
        lsCtaContDebeBME = Left(lsCtaContDebeB, 2) & "2" & Right(lsCtaContDebeB, Len(lsCtaContDebeB) - 3)
        lsCtaContDebeSME = Left(lsCtaContDebeS, 2) & "2" & Right(lsCtaContDebeS, Len(lsCtaContDebeS) - 3)
        lsCtaContDebeRHME = Left(lsCtaContDebeRH, 2) & "2" & Right(lsCtaContDebeRH, Len(lsCtaContDebeRH) - 3) 'RRHH Descuento Planillas
        lsCtaContDebeRHJME = Left(lsCtaContDebeRHJ, 2) & "2" & Right(lsCtaContDebeRHJ, Len(lsCtaContDebeRHJ) - 3) 'RRHH Judicial
        lsCtaContDebeSeguME = Left(lsCtaContDebeSegu, 2) & "2" & Right(lsCtaContDebeSegu, Len(lsCtaContDebeSegu) - 3) 'Seguros
        lsCtaContDebePagoVariosME = Left(lsCtaContDebePagoVarios, 2) & "2" & Right(lsCtaContDebePagoVarios, Len(lsCtaContDebePagoVarios) - 3) 'Pagos Varios
    End If
End If

'*************************************NAGL INC1712260008**********************************************
lsProveedorCtasApertSBS = oDCaja.GetProveedoresCtasAperturadasSBS(gsOpeCod)
If lsProveedorCtasApertSBS <> "" Then
    Set rs = oDCaja.GetDatosProvisionesProveedores("'" & lsProveedorLeasingSMN & "','" & lsProveedorLeasingSME & "','" & lsProveedorLeasingMN & "','" & lsProveedorLeasingME & "','" & lsCtaContDebeB & "','" & lsCtaContDebeS & "','" & lsCtaContDebeRH & "','" & lsCtaContDebeRHJ & "', '" & lsCtaContDebeSegu & "', '" & lsCtaContDebePagoVarios & "','" & lsProveedorFEPCMACMN & "','" & lsProveedorFEPCMACME & "'," & lsProveedorCtasApertSBS & "", lsDocs, txtFechaDel, txtFechaAl, , 3, cCtaDetraTemp, "'" & lsCtaContDebeBME & "','" & lsCtaContDebeSME & "','" & lsCtaContDebeRHME & "','" & lsCtaContDebeRHJME & "', '" & lsCtaContDebeSeguME & "', '" & lsCtaContDebePagoVariosME & "'", lsTipoInterfaz, Me.txtProveedor.Text, psTipo, False, Me.txtAge.Text, lsTipoB, lnTipoPago)
Else
    Set rs = oDCaja.GetDatosProvisionesProveedores("'" & lsProveedorLeasingSMN & "','" & lsProveedorLeasingSME & "','" & lsProveedorLeasingMN & "','" & lsProveedorLeasingME & "','" & lsCtaContDebeB & "','" & lsCtaContDebeS & "','" & lsCtaContDebeRH & "','" & lsCtaContDebeRHJ & "', '" & lsCtaContDebeSegu & "', '" & lsCtaContDebePagoVarios & "','" & lsProveedorFEPCMACMN & "','" & lsProveedorFEPCMACME & "'", lsDocs, txtFechaDel, txtFechaAl, , 3, cCtaDetraTemp, "'" & lsCtaContDebeBME & "','" & lsCtaContDebeSME & "','" & lsCtaContDebeRHME & "','" & lsCtaContDebeRHJME & "', '" & lsCtaContDebeSeguME & "', '" & lsCtaContDebePagoVariosME & "'", lsTipoInterfaz, Me.txtProveedor.Text, psTipo, False, Me.txtAge.Text, lsTipoB, lnTipoPago)
End If
'***END NAGL Agregó lsProveedorCtasApertSBS, en este método según INC1712260008 y Condicional**********


'If gbPermisoLogProveedorAG Then
'    Set rs = oDCaja.GetDatosProvisionesProveedores("'" & lsCtaContDebeB & "','" & lsCtaContDebeS & "'", lsDocs, txtFechaDel, txtFechaAl, , 3, cCtaDetraTemp, "'" & lsCtaContDebeBME & "','" & lsCtaContDebeSME & "'", lsTipoInterfaz, Me.txtProveedor.Text, psTipo, True, gsCodAge)
'Else
    'Set rs = oDCaja.GetDatosProvisionesProveedores("'" & lsProveedorLeasingSMN & "','" & lsProveedorLeasingSME & "','" & lsProveedorLeasingMN & "','" & lsProveedorLeasingME & "','" & lsCtaContDebeB & "','" & lsCtaContDebeS & "','" & lsCtaContDebeRH & "','" & lsCtaContDebeRHJ & "', '" & lsCtaContDebeSegu & "', '" & lsCtaContDebePagoVarios & "','" & lsProveedorFEPCMACMN & "','" & lsProveedorFEPCMACME & "' ", lsDocs, txtFechaDel, txtFechaAl, , 3, cCtaDetraTemp, "'" & lsCtaContDebeBME & "','" & lsCtaContDebeSME & "','" & lsCtaContDebeRHME & "','" & lsCtaContDebeRHJME & "', '" & lsCtaContDebeSeguME & "', '" & lsCtaContDebePagoVariosME & "'", lsTipoInterfaz, Me.txtProveedor.Text, psTipo, False, Me.txtAge.Text, lsTipoB, lnTipoPago)'Comentado by NAGL según INC1712260008
    
'End If

Set oDCaja = Nothing

If rs.EOF Then
   RSClose rs
   cmdProcesar.Enabled = True
   MsgBox "No existen Comprobantes Pendientes", vbInformation, "Aviso"
   Exit Sub
End If

nCantTempo = 0

Do While Not rs.EOF
        fg.AdicionaFila
        nItem = fg.Row
        
        fg.TextMatrix(nItem, 1) = nItem
        fg.TextMatrix(nItem, 3) = Mid(rs!cDocAbrev & space(3), 1, 3) & " " & rs!cDocNro
        fg.TextMatrix(nItem, 4) = rs!dDocFecha
        fg.TextMatrix(nItem, 5) = PstaNombre(rs!cPersona, True)
        fg.TextMatrix(nItem, 6) = rs!cMovDesc
        fg.TextMatrix(nItem, 7) = Format(rs!nMovImporte, gsFormatoNumeroView)
        fg.TextMatrix(nItem, 8) = rs!cPersCod
        fg.TextMatrix(nItem, 9) = rs!cMovNro
        fg.TextMatrix(nItem, 10) = rs!nMovNro
        fg.TextMatrix(nItem, 11) = rs!nDocTpo
        fg.TextMatrix(nItem, 12) = rs!cDocNro
        fg.TextMatrix(nItem, 13) = rs!cCtaContCod
        fg.TextMatrix(nItem, 14) = GetFechaMov(rs!cMovNro, True)
        
        'fg.TextMatrix(nItem, 23) = 'Format(rs!nPenalidad, gsFormatoNumeroView)
        If gsOpeCod <> OpeCGOpeProvRechazo And Me.fraBuscar.Visible = True Then
            fg.TextMatrix(nItem, 24) = rs!Movenvio
            fg.TextMatrix(nItem, 25) = rs!Agencia
        ElseIf gsOpeCod = OpeCGOpeProvPagoListSUNAT Then
            fg.TextMatrix(nItem, 24) = rs!Movenvio
            fg.TextMatrix(nItem, 25) = rs!Agencia
        End If
        'ALPA 20090317**********************
        fg.TextMatrix(nItem, 26) = rs!nPenalidad
        '************************************
'        If rs!nPenalidad > 0 Then
'            fg.Col = 23
'            fg.Row = nItem
'            fg.ForeColorRow (vbBlue)
'        End If
        If lsTipoInterfaz = "lbReporte" Then
            If rs!nMovImporteSoles = rs!nMovImporte Then
                fg.TextMatrix(nItem, 15) = Format(rs!nMovImporteSoles - rs!MontoPagadoSUNAT, gsFormatoNumeroView)
            Else
                fg.TextMatrix(nItem, 15) = "0.00"
            End If
            fg.TextMatrix(nItem, 21) = Format(rs!MontoPagadoSUNAT, gsFormatoNumeroView)
            fg.TextMatrix(nItem, 22) = Format(rs!MontoPagadoSUNATS, gsFormatoNumeroView)
        End If
        
        If lsTipoInterfaz = "PAGO" And Me.optTipo(1) = True Then
            fg.TextMatrix(nItem, 17) = Format(rs!nimportecoactivo, gsFormatoNumeroView)
            'Se comento la linea para que se puede pagar indistinatmente
            'los proveedores con pago sunat, y se agrego la linea siguiente. GITU
            'fg.TextMatrix(nItem, 18) = Format(rs!montopago - rs!MontoPagadoSUNAT, gsFormatoNumeroView) '- rs!nPenalidad
            fg.TextMatrix(nItem, 18) = Format(rs!montopago, gsFormatoNumeroView) '- rs!nPenalidad
            If rs!nMovImporteSoles <> rs!nMovImporte Then
                fg.TextMatrix(nItem, 15) = Format(Round((rs!nMovImporte - rs!MontoPagadoSUNAT) * rs!nTpoCambio, 2), gsFormatoNumeroView)
                fg.TextMatrix(nItem, 20) = Format(Round((rs!nMovImporte - rs!MontoPagadoSUNAT) * lnTCFijo, 2) - Round(CDbl(fg.TextMatrix(nItem, 18)) * lnTCFijo, 2) - Round(Round(rs!nimportecoactivo / rs!nTpoCambio, 2) * lnTCFijo, 2), "0.00")
                fg.TextMatrix(nItem, 21) = Format(rs!MontoPagadoSUNAT, gsFormatoNumeroView)
                fg.TextMatrix(nItem, 22) = Format(rs!MontoPagadoSUNATS, gsFormatoNumeroView)
                
            Else
                fg.TextMatrix(nItem, 15) = Format(rs!nMovImporteSoles - rs!MontoPagadoSUNAT, gsFormatoNumeroView)
                fg.TextMatrix(nItem, 20) = "0.00"
            End If
        ElseIf lsTipoInterfaz = "PAGO" And Me.optTipo(1) = False Then
            fg.TextMatrix(nItem, 15) = Format(rs!nMovImporteSoles, gsFormatoNumeroView)
        End If
        
        If (lsTipoInterfaz = "PAGO" Or lsTipoInterfaz = "lbReporte") Then
           fg.TextMatrix(nItem, 19) = IIf(rs!nMovImporteSoles = rs!nMovImporte, "SOLES", "DOLARES")
        End If
        fg.TextMatrix(nItem, 16) = IIf(IsNull(rs!cPersIdNro), "RUC NO REGISTRADO", Trim(rs!cPersIdNro))
     
    rs.MoveNext
Loop
RSClose rs
fg.Row = 1
txtMovDesc = fg.TextMatrix(1, 6)

If nCantTempo > 0 Then
    MsgBox "Existe(n) " & nCantTempo & " registro(s) que no se cargó porque aún falta registrar la detracción", vbInformation, "Aviso"
Else
    If lsTipoInterfaz = "lbReporte" Then
        Me.cmdGenerarArch.Enabled = True
    End If
End If

Exit Sub
ErrCargaProveedores:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub mskFiltro_GotFocus()
    mskFiltro.SelStart = 0
    mskFiltro.SelLength = 50
End Sub

Private Sub mskFiltro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
    End If
End Sub

Private Sub optAge_Click(Index As Integer)
    Select Case Index
        Case 0
            Me.txtAge.Text = ""
            Me.lblAgencia.Caption = ""
            Me.txtAge.Enabled = False
            lsTipoB = "T"
        Case 1
            Me.txtAge.Enabled = True
            lsTipoB = "A"
        Case 2
            Me.txtAge.Text = ""
            Me.lblAgencia.Caption = ""
            Me.txtAge.Enabled = False
            lsTipoB = "L"
    End Select
End Sub

Private Sub optTipo_Click(Index As Integer)
    Select Case Index
        Case 0
            Me.fraFecha.Visible = True
            Me.fraBuscar.Visible = False
            Me.FraTipoB.Visible = False
            Me.fraFecha.Height = 680
            
            Me.Label3.Left = 2670
            Me.Label3.Top = 330
            Me.txtFechaDel.Left = 3390
            Me.txtFechaDel.Top = 240
            
            Me.Label4.Left = 4620
            Me.Label4.Top = 330
            Me.txtFechaAl.Top = 240
            Me.txtFechaAl.Left = 5190
            
            Me.cmdProcesar.Top = 630
            Me.cmdProcesar.Left = 7720
        Case 1
        
            Me.fraFecha.Visible = False
            Me.fraBuscar.Visible = True
            Me.cmdProcesar.Left = 9960
            
            Me.Label3.Left = 5700
            Me.Label3.Top = 390
            Me.txtFechaDel.Left = 5700
            Me.txtFechaDel.Top = 630
            
            Me.Label4.Left = 6990
            Me.Label4.Top = 390
            Me.txtFechaAl.Top = 630
            Me.txtFechaAl.Left = 5190
            
            Me.cmdProcesar.Top = 540
            Me.cmdProcesar.Left = 9930
    End Select
End Sub

Private Sub txtAge_EmiteDatos()
    Me.lblAgencia = Me.txtAge.psDescripcion
End Sub

Private Sub txtBuscaEntidad_EmiteDatos()
Dim oCtaIf As NCajaCtaIF
Set oCtaIf = New NCajaCtaIF
lblCtaDesc = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaEntidad, 18, 10)) + " " + txtBuscaEntidad.psDescripcion
If txtBuscaEntidad <> "" Then
   cmdDoc.SetFocus
End If
Set oCtaIf = Nothing
End Sub

Private Sub txtFechaAl_GotFocus()
    fEnfoque txtFechaAl
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Not ValFecha(txtFechaAl) Then Exit Sub
   cmdProcesar.SetFocus
End If
End Sub

Private Sub txtFechaDel_GotFocus()
fEnfoque txtFechaDel
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Not ValFecha(txtFechaDel) Then Exit Sub
   txtFechaAl.SetFocus
End If
End Sub



Private Sub txtProveedor_EmiteDatos()
    If Me.txtProveedor <> "" Then
        Me.cmdDoc.Enabled = False
        Me.CmdImprimir.Enabled = False
        Me.cmdProcesar.Enabled = False
        DoEvents
        CargaProveedores
        Me.cmdDoc.Enabled = True
        Me.CmdImprimir.Enabled = True
        Me.cmdProcesar.Enabled = True
    End If
End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGenerarArch.SetFocus
    Else
        KeyAscii = NumerosDecimales(Me.txtTipoCambio, KeyAscii, 8, 3)
    End If
End Sub

Private Sub AplicaTipoCambio()
    Dim i As Integer
    For i = 1 To Me.fg.Rows - 1
        If fg.TextMatrix(i, 19) <> "SOLES" Then
            fg.TextMatrix(i, 15) = Format(Round((CDbl(fg.TextMatrix(i, 7)) - CDbl(fg.TextMatrix(i, 21))) * CDbl(Me.txtTipoCambio), 2), gsFormatoNumeroView)
        End If
    Next i
End Sub

Private Sub vFormPago_MenuItemClick(MenuNumber As Long, MenuItem As Long)
txtFechaMov.Enabled = False
txtFechaMov = gdFecSis
Select Case MenuItem
    Case 1: 'Efectivo
            fraEntidad.Visible = False
            fraEntidad.Visible = False
            lsDocTpo = "-1"
            'ALPA 20090318***************************
            lnTipoPago = 0
            '****************************************
            cmdDoc_Click
    Case 2: ' TpoDocCarta
            fraEntidad.Visible = True
            lsDocTpo = TpoDocCarta
            txtFechaMov.Enabled = True
            'ALPA 20090318***************************
            lnTipoPago = 0
            '****************************************
   Case 3:  ' TpoDocOrdenPago
            fraEntidad.Visible = False
            lsDocTpo = TpoDocOrdenPago
            'ALPA 20090318***************************
            lnTipoPago = 0
            '****************************************
            cmdDoc_Click
   Case 4:  'Cheque
            fraEntidad.Visible = True
            lsDocTpo = TpoDocCheque
            txtFechaMov.Enabled = True
            'ALPA 20090318***************************
            lnTipoPago = 0
            '****************************************
   Case 5:  'Nota de Abono
            fraEntidad.Visible = False
            lsDocTpo = TpoDocNotaAbono
            'ALPA 20090318***************************
            lnTipoPago = 0
            '****************************************
            cmdDoc_Click
    'ALPA 20090318***************************
    Case 6:  'Penalidad
            fraEntidad.Visible = False
            fraEntidad.Visible = False
            lsDocTpo = TpoDocRetenciones
            lnTipoPago = 1
            cmdProcesar_Click
    '****************************************
End Select
If fraEntidad.Visible Then
   txtBuscaEntidad.SetFocus
Else
   If cmdDoc.Visible Then
    cmdDoc.SetFocus
   End If
End If
End Sub

Private Sub vFormPagoAge_MenuItemClick(MenuNumber As Long, MenuItem As Long)
    txtFechaMov.Enabled = False
    txtFechaMov = gdFecSis
    Select Case MenuItem
       Case 1:  ' TpoDocOrdenPago
                fraEntidad.Visible = False
                lsDocTpo = TpoDocOrdenPago
                cmdDoc_Click
    End Select
    If fraEntidad.Visible Then
       txtBuscaEntidad.SetFocus
    Else
       If cmdDoc.Visible Then
        cmdDoc.SetFocus
       End If
    End If
End Sub

