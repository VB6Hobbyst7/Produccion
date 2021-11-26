VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDocPago 
   Caption         =   "Documento de Pago"
   ClientHeight    =   6045
   ClientLeft      =   1080
   ClientTop       =   1905
   ClientWidth     =   9465
   Icon            =   "frmDocPago.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9465
   Begin VB.Frame Frame7 
      Height          =   990
      Left            =   7200
      TabIndex        =   41
      Top             =   120
      Width           =   2145
      Begin MSMask.MaskEdBox txtDocFecha 
         Height          =   345
         Left            =   525
         TabIndex        =   43
         Top             =   480
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha de Documento"
         Height          =   195
         Left            =   285
         TabIndex        =   42
         Top             =   210
         Width           =   1635
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Glosa "
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
      Height          =   975
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   7005
      Begin VB.TextBox txtGlosa 
         CausesValidation=   0   'False
         Height          =   615
         Left            =   180
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   240
         Width           =   6645
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   5580
      Width           =   1260
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6795
      TabIndex        =   11
      Top             =   5580
      Width           =   1260
   End
   Begin TabDlg.SSTab pageDoc 
      CausesValidation=   0   'False
      Height          =   4275
      Left            =   150
      TabIndex        =   13
      Top             =   1260
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   7541
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   5
      TabHeight       =   564
      TabCaption(0)   =   "Car&ta   "
      TabPicture(0)   =   "frmDocPago.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "rtfDoc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdPlantilla"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "C&heque   "
      TabPicture(1)   =   "frmDocPago.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Shape1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "&Orden de Pago  "
      TabPicture(2)   =   "frmDocPago.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(2)=   "Shape2"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame6 
         Enabled         =   0   'False
         Height          =   675
         Left            =   -74895
         TabIndex        =   39
         Top             =   3300
         Width           =   3105
         Begin VB.TextBox txtVoucherOP 
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
            Left            =   1290
            TabIndex        =   10
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label Label8 
            Caption         =   "Nro. Voucher"
            Height          =   285
            Left            =   150
            TabIndex        =   40
            Top             =   270
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74880
         TabIndex        =   25
         Top             =   480
         Width           =   8895
         Begin VB.TextBox txtNroOP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   315
            Left            =   5100
            MaxLength       =   20
            TabIndex        =   7
            Top             =   720
            Width           =   1980
         End
         Begin VB.TextBox txtFechaOP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5100
            TabIndex        =   27
            Top             =   300
            Width           =   1155
         End
         Begin VB.TextBox txtImpOP 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6840
            TabIndex        =   26
            Top             =   300
            Width           =   1755
         End
         Begin VB.TextBox txtPersOP 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1380
            TabIndex        =   8
            Top             =   1560
            Width           =   7275
         End
         Begin VB.TextBox txtCantOP 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1380
            TabIndex        =   9
            Top             =   1980
            Width           =   7275
         End
         Begin VB.PictureBox Animation1 
            BackColor       =   &H00E0E0E0&
            Height          =   705
            Left            =   270
            ScaleHeight     =   645
            ScaleWidth      =   2265
            TabIndex        =   36
            Top             =   270
            Width           =   2325
         End
         Begin VB.Label lblSimbOP 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   6240
            TabIndex        =   35
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Orden de Pago"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3600
            TabIndex        =   31
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "La suma de"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   2040
            Width           =   1155
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ica ,"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   4440
            TabIndex        =   29
            Top             =   360
            Width           =   435
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Pague a la orden  de :"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   120
            TabIndex        =   28
            Top             =   1440
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdPlantilla 
         Caption         =   "&Plantilla de Carta"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   3750
         Width           =   1995
      End
      Begin VB.Frame Frame1 
         Height          =   675
         Left            =   6660
         TabIndex        =   19
         Top             =   360
         Width           =   2355
         Begin VB.TextBox txtNroCarta 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   915
            MaxLength       =   8
            TabIndex        =   1
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nro Carta"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   300
            Width           =   855
         End
      End
      Begin RichTextLib.RichTextBox rtfDoc 
         Height          =   3255
         Left            =   120
         TabIndex        =   2
         Top             =   420
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   5741
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmDocPago.frx":035E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   8895
         Begin VB.CommandButton btnNroCH 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6560
            TabIndex        =   45
            Top             =   720
            Width           =   350
         End
         Begin VB.TextBox txtNroCH 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5100
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   740
            Width           =   1455
         End
         Begin VB.TextBox txtCantCH 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1380
            TabIndex        =   5
            Top             =   1980
            Width           =   7275
         End
         Begin VB.TextBox txtPersCH 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1380
            TabIndex        =   4
            Top             =   1560
            Width           =   7275
         End
         Begin VB.TextBox txtImpCH 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6840
            TabIndex        =   17
            Top             =   300
            Width           =   1755
         End
         Begin VB.TextBox txtFechaCH 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5100
            TabIndex        =   16
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label lblCuentaCH 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   3255
         End
         Begin VB.Label lblBancoCH 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   12
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   24
            Top             =   180
            Width           =   3420
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Paguese a la orden de"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   120
            TabIndex        =   23
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lblSimbCH 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   6300
            TabIndex        =   22
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4110
            TabIndex        =   21
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "La suma de"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   2040
            Width           =   1155
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro Cheque"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3900
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   675
         Left            =   -74895
         TabIndex        =   37
         Top             =   3300
         Width           =   3105
         Begin VB.TextBox txtVoucherCh 
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
            Left            =   1320
            TabIndex        =   6
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label Label6 
            Caption         =   "Nro. Voucher"
            Height          =   285
            Left            =   150
            TabIndex        =   38
            Top             =   270
            Width           =   1065
         End
      End
      Begin VB.Shape Shape2 
         Height          =   2805
         Left            =   -74895
         Top             =   465
         Width           =   8925
      End
      Begin VB.Shape Shape1 
         Height          =   2805
         Left            =   -74895
         Top             =   465
         Width           =   8925
      End
   End
   Begin RichTextLib.RichTextBox rtfIMP 
      Height          =   375
      Left            =   180
      TabIndex        =   32
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"frmDocPago.frx":03DE
   End
End
Attribute VB_Name = "frmDocPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim txtPlantilla As String
Dim Pag0 As Boolean, Pag1 As Boolean, Pag2 As Boolean

Public vsGlosa As String
Public vdFechaDoc As Date
Public vbOk As Boolean
Public vsTpoDoc As TpoDoc
Public vsNroDoc As String
Public vnImporte As Currency
Public vsPersCod As String
Public vsPersNombre As String
Public vsNroVoucher As String
Public vsFileCarta As String
Public vsTipoDocVoucher As String
Public vsFormaDoc As String
Public lnMagIzq As Integer
Public lnMagDer As Integer
Public lnMagSup As Integer

Private vsCodAge As String

Public lsPlantillaDoc As String

Dim lsOpeDesc As String
Dim lsOpeCod As String
Dim ldFechaSist As Date
Dim lsEntiOrig  As String
Dim lsCtaEntiOrig  As String
Dim lsEntiDest As String
Dim lsCtaEntiDest As String
Dim lsMovNro As String
Dim lsEmpresaRuc As String
Dim lsSubCtaIF As String

Dim lsPersIfCod As String
Dim lsPersNombre As String
Dim lnColPage As Integer
Dim lsSimbolo As String
Public lbIngresoPers As Boolean

Dim oNContFun As NContFunciones
Dim oNPlant As NPlantilla
Dim oDocPago As clsDocPago
Dim lnDocTpo As TpoDoc
Dim lbModificaMonto As Boolean

'EJVG20121130 ***
Dim fsIFTpo As String
Dim fsPersCodIF As String
Dim fsCtaIFCod As String
'END EJVG *******

'Public Sub InicioCheque(psOpeCod As String, ByVal psOpeDesc As String, psPersNombre As String, psGlosa As String, pnImporte As Currency, pdFechaSist As Date, psEmpresaRuc As String, psSubCtaIF As String, ByVal psEntidadOrig As String, ByVal psCtaEntidadOrig As String, Optional psDocNroVoucher As String = "", Optional pbIngPers As Boolean, Optional psCodAge As String = "", Optional pbModificaMonto As Boolean = False)
Public Sub InicioCheque(psOpeCod As String, ByVal psOpeDesc As String, psPersNombre As String, psGlosa As String, pnImporte As Currency, pdFechaSist As Date, psEmpresaRuc As String, psSubCtaIF As String, ByVal psEntidadOrig As String, ByVal psCtaEntidadOrig As String, Optional psDocNroVoucher As String = "", Optional pbIngPers As Boolean, Optional psCodAge As String = "", Optional pbModificaMonto As Boolean = False, Optional ByVal psIFTpo As String, Optional ByVal psPersCod As String, Optional ByVal psCtaIFCod As String) 'EJVG20121129
Pag0 = False
Pag1 = True
Pag2 = False

lbIngresoPers = pbIngPers
lsOpeCod = psOpeCod
lsEntiOrig = psEntidadOrig
lsOpeDesc = psOpeDesc
lsCtaEntiOrig = psCtaEntidadOrig
ldFechaSist = pdFechaSist
lsEmpresaRuc = psEmpresaRuc
lsSubCtaIF = psSubCtaIF
vsPersNombre = psPersNombre
vsNroVoucher = psDocNroVoucher
vnImporte = pnImporte
vsGlosa = psGlosa
vsCodAge = psCodAge
lbModificaMonto = pbModificaMonto
'EJVG20121129 ***
fsIFTpo = psIFTpo
fsPersCodIF = psPersCod
fsCtaIFCod = psCtaIFCod
'END EJVG *******
Me.Show 1
End Sub
Public Sub InicioCarta(psOpeCod As String, ByVal psOpeDesc As String, psGlosa As String, psFile As String, _
                        pnImporte As Currency, pdFechaSist As Date, _
                        psEntiOrig As String, psCtaEntiOrig As String, _
                        psEntiDest As String, psCtaEntiDest As String, _
                        psMovNro As String, Optional psDocNroVoucher As String = "")

Pag0 = True
Pag1 = False
Pag2 = False
lsOpeCod = psOpeCod
ldFechaSist = pdFechaSist
lsEntiOrig = Trim(psEntiOrig)
lsCtaEntiOrig = Trim(psCtaEntiOrig)
lsEntiDest = Trim(psEntiDest)
lsCtaEntiDest = Trim(psCtaEntiDest)
lsMovNro = psMovNro
lsOpeDesc = Trim(psOpeDesc)
vsFileCarta = psFile
vnImporte = pnImporte
vsGlosa = psGlosa
vdFechaDoc = pdFechaSist
lnColPage = 80
Me.Show 1
End Sub
Public Sub InicioOrdenPago(ByVal psOpeCod As String, ByVal psOpeDesc As String, psPersNombre As String, psGlosa As String, pnImporte As Currency, pdFechaSist As Date, Optional psDocNroVoucher As String = "", Optional pbIngPers As Boolean, Optional psCodAge As String = "", Optional pbModificaMonto As Boolean = False)
Pag0 = False
Pag1 = False
Pag2 = True
lbIngresoPers = pbIngPers
lsOpeCod = psOpeCod
ldFechaSist = pdFechaSist
lsOpeDesc = psOpeDesc
vnImporte = pnImporte
vsGlosa = psGlosa
vsPersNombre = psPersNombre
vsNroVoucher = psDocNroVoucher
vsCodAge = psCodAge
lbModificaMonto = pbModificaMonto
Me.Show 1
End Sub
'EJVG20120113 ***
Private Sub btnNroCH_Click()
    Dim NroCheque As String
    NroCheque = frmSeleccionaCheque.Inicio(fsIFTpo, fsPersCodIF, fsCtaIFCod)
    Me.txtNroCH.Text = NroCheque
End Sub
'END EJVG *******
Private Sub cmdAceptar_Click()
Dim oConect As DConecta
Dim sSql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim oDoc As DDocumento

Set oDoc = New DDocumento
Set oConect = New DConecta
If oConect.AbreConexion = False Then Exit Sub
If Val(txtImpOP) = 0 Or Val(txtImpCH) = 0 Then
    MsgBox "Cantidad no puede ser igual a cero", vbInformation, "Aviso"
    Exit Sub
End If
If ValFecha(txtDocFecha) = False Then Exit Sub

vsGlosa = txtGlosa.Text
vdFechaDoc = txtDocFecha.Text
Select Case pageDoc.Tab
     Case 0
          If txtNroCarta = "" Then
            MsgBox "Falta Indicar Nro. de Documento", vbInformation, "Aviso"
            Exit Sub
          End If
          If Len(Trim(rtfDoc.Text)) = 0 Then
            MsgBox "Carta no ha sido seleccionada - Ingresada - Digitada", vbInformation, "Aviso"
            cmdPlantilla.SetFocus
            Exit Sub
          End If
          vsTpoDoc = Format(TpoDocCarta, "00")
          vsNroDoc = txtNroCarta.Text
          vsFormaDoc = oDocPago.ProcesaPlantilla(txtPlantilla, True, lsMovNro, ldFechaSist, _
                        lsEntiOrig, lsEntiDest, vnImporte, lsSimbolo, lsCtaEntiOrig, lsCtaEntiDest, _
                        txtNroCarta, lnColPage, lnMagIzq, lnMagDer, Mid(lsOpeCod, 3, 1), lnMagSup)
     Case 1
          If txtNroCH = "" Then
            MsgBox "Falta indicar Nro. de Documento", vbInformation, "Aviso"
            Exit Sub
          End If
          If txtVoucherCh = "" Then
             MsgBox "Falta Ingresar Nro. de Voucher ", vbInformation, "Aviso"
             Exit Sub
          End If
          If ValFecha(txtFechaCH) = False Then Exit Sub
          vdFechaDoc = CDate(txtFechaCH)
          vsTpoDoc = Format(TpoDocCheque, "00")
          vsNroDoc = txtNroCH.Text
          vsPersNombre = txtPersCH
          vnImporte = txtImpCH
          vsNroVoucher = txtVoucherCh
          vsFormaDoc = FormaCheque(lsSubCtaIF, lsEmpresaRuc, vsPersNombre, vdFechaDoc, vnImporte, Mid(lsOpeCod, 3, 1))
     Case 2
          If txtNroOP = "" Then
            MsgBox "Falta indicar Nro. de Documento", vbInformation, "Aviso"
            Exit Sub
          End If
          If txtVoucherOP = "" Then
             MsgBox "Falta Ingresar Nro. de Voucher", vbInformation, "Aviso"
             Exit Sub
          End If
          If ValFecha(txtFechaOP) = False Then Exit Sub
          vdFechaDoc = CDate(txtFechaOP)
          vsTpoDoc = Format(TpoDocOrdenPago, "00")
          vsNroDoc = Trim(txtNroOP.Text)
          vsPersNombre = Trim(txtPersOP)
          vnImporte = txtImpOP
          vsNroVoucher = Trim(txtVoucherOP)
          'vsFormaDoc = FormaOrdenPago(vsPersNombre, vdFechaDoc, vnImporte, Mid(lsOpeCod, 3, 1))
          vsFormaDoc = FormaOrdenPagoTalon(vsPersNombre, vdFechaDoc, vnImporte, Mid(lsOpeCod, 3, 1))
        End Select
        If pageDoc.Tab <> 1 Then 'EJVG20121130
        If oDoc.VerificaDoc(vsTpoDoc, vsNroDoc, vsPersCod) = True Then
            MsgBox "Documento Ingresado ya fue registrado", vbInformation, "Aviso"
            Exit Sub
        End If
        Else
            If oDoc.VerificaChequeCMAC(fsIFTpo, fsPersCodIF, fsCtaIFCod, vsNroDoc) Then
                MsgBox "Cheque Ingresado ya se esta usando", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
 vbOk = True
Unload Me
End Sub
Private Sub cmdCancelar_Click()
vbOk = False
Unload Me
End Sub
Private Sub cmdPlantilla_Click()
frmPlantillas.lnMagIzq = lnMagIzq
frmPlantillas.lnMagDer = lnMagDer
frmPlantillas.lnMagSup = lnMagSup

frmPlantillas.Inicio lsPlantillaDoc, lsOpeCod
lsPlantillaDoc = frmPlantillas.lsPlantillaID
lnMagIzq = frmPlantillas.lnMagIzq
lnMagDer = frmPlantillas.lnMagDer
lnMagSup = frmPlantillas.lnMagSup
txtPlantilla = oNPlant.GetPlantillaDoc(lsPlantillaDoc)
rtfDoc.Text = oDocPago.ProcesaPlantilla(txtPlantilla, False, lsMovNro, vdFechaDoc, lsEntiOrig, lsEntiDest, vnImporte, lsSimbolo, lsCtaEntiOrig, lsCtaEntiDest, vsNroDoc, gnColPage, lnMagIzq, lnMagDer, Mid(lsOpeCod, 3, 1), lnMagSup)
txtNroCarta.SetFocus
End Sub

Private Sub Form_Load()
Set oDocPago = New clsDocPago

Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2, Me.Width, Me.Height
Set oNContFun = New NContFunciones
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing


vbOk = False
pageDoc.Tab = 0
Set oNPlant = New NPlantilla
Me.Caption = lsOpeDesc & ": Documento de Pago"
If Pag2 Then
   'If Dir(App.Path & "\videos\logoorden.avi") <> "" Then
   '   Animation1.AutoPlay = True
   '   Animation1.Open App.Path & "\videos\logoorden.avi"
   'End If
   txtNroOP = oNContFun.GeneraDocNro(TpoDocOrdenPago, , Mid(lsOpeCod, 3, 1), Right(vsCodAge, 2))
    lnDocTpo = TpoDocOrdenPago
End If

If vnImporte > 0 Then
    txtCantCH.Text = ConvNumLet(vnImporte, Mid(lsOpeCod, 3, 1))
    txtCantOP.Text = ConvNumLet(vnImporte, Mid(lsOpeCod, 3, 1))
End If

Dim oConst As NConstSistemas
Set oConst = New NConstSistemas
If lnMagDer = 0 Then
    lnMagDer = oConst.LeeConstSistema(gConstSistMargenDerCartas)   'nContFunc.GetValorMargenes(MargenDerecho)
End If
If lnMagIzq = 0 Then
    lnMagIzq = oConst.LeeConstSistema(gConstSistMagenIzqCartas)     'nContFunc.GetValorMargenes(MargenIzquierdo)
End If
If lnMagSup = 0 Then
    lnMagSup = oConst.LeeConstSistema(gConstSistMargenSupCartas)    'nContFunc.GetValorMargenes(MargenSuperior)
End If
Set oConst = Nothing

If vsGlosa <> "" Then
   txtGlosa.Text = vsGlosa
End If
If vsNroDoc <> "" Then
   txtNroCarta.Text = vsNroDoc
   txtNroCH.Text = vsNroDoc
End If

txtFechaCH.Text = Format(ldFechaSist, "dd/mm/yyyy")
txtFechaOP.Text = Format(ldFechaSist, "dd/mm/yyyy")
txtImpCH = Format(vnImporte, "####,####0.00")
txtImpOP = Format(vnImporte, "#####,###0.00")

lsSimbolo = IIf(Mid(lsOpeCod, 3, 1) = gMonedaNacional, gcPEN_SIMBOLO, "US$") 'marg ers044-2016

txtPersCH.Text = vsPersNombre
txtPersOP.Text = vsPersNombre

lblBancoCH.Caption = lsEntiOrig
lblCuentaCH.Caption = lsCtaEntiOrig

pageDoc.TabVisible(0) = Pag0
pageDoc.TabVisible(1) = Pag1
pageDoc.TabVisible(2) = Pag2
txtPlantilla = ""

If Pag0 Then
    If lsPlantillaDoc = "" Then
        lsPlantillaDoc = oNPlant.GetNroPlantillaOpe(lsOpeCod)
    End If
    If Not lsPlantillaDoc = "" Then
        txtPlantilla = oNPlant.GetPlantillaDoc(lsPlantillaDoc)
    End If
    If Len(Trim(txtPlantilla)) = 0 Then
        MsgBox "No Existe plantilla para el documento de esta operación...", vbInformation, "Aviso"
    Else
        rtfDoc.Text = oDocPago.ProcesaPlantilla(txtPlantilla, False, lsMovNro, vdFechaDoc, lsEntiOrig, lsEntiDest, vnImporte, lsSimbolo, lsCtaEntiOrig, lsCtaEntiDest, vsNroDoc, gnColPage, lnMagIzq, lnMagDer, Mid(lsOpeCod, 3, 1), lnMagSup)
    End If
    txtDocFecha.Enabled = True
    lnDocTpo = TpoDocCarta
End If
If Pag1 Then
   txtPersCH.Enabled = lbIngresoPers
   txtVoucherCh = vsNroVoucher
   lnDocTpo = TpoDocCheque
   txtImpCH.Enabled = lbModificaMonto
End If
If Pag2 Then
   txtPersOP.Enabled = lbIngresoPers
   txtVoucherOP = vsNroVoucher
   lnDocTpo = TpoDocOrdenPago
   txtImpOP.Enabled = lbModificaMonto
End If
txtDocFecha.Text = ldFechaSist
ActivaCampos
Screen.MousePointer = 0


End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oNContFun = Nothing
End Sub

Private Sub pageDoc_Click(PreviousTab As Integer)
ActivaCampos
End Sub
Private Sub ActivaCampos()
cmdAceptar.Enabled = False
Select Case pageDoc.Tab
       Case 0
            txtNroCarta.Enabled = True
            cmdPlantilla.Enabled = True
            rtfDoc.Enabled = True
            txtNroCH.Enabled = False
            txtNroOP.Enabled = False
       Case 1
            txtNroCarta.Enabled = False
            cmdPlantilla.Enabled = False
            rtfDoc.Enabled = False
            txtNroCH.Enabled = True
            txtNroOP.Enabled = False
       Case 2
            txtNroCarta.Enabled = False
            cmdPlantilla.Enabled = False
            rtfDoc.Enabled = False
            txtNroCH.Enabled = False
            txtNroOP.Enabled = True
End Select
End Sub

Private Sub pageDoc_GotFocus()
Select Case pageDoc.Tab
    Case 0:    txtNroCarta.SetFocus
    Case 1:    txtNroCH.SetFocus
    Case 2:    txtNroOP.SetFocus
End Select
End Sub


Private Sub txtCantCH_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdAceptar.Enabled = True
   cmdAceptar.SetFocus
End If
End Sub

Private Sub txtCantOP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdAceptar.Enabled = True
   cmdAceptar.SetFocus
End If
End Sub

Private Sub txtDocFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   vdFechaDoc = txtDocFecha.Text
   rtfDoc.Text = oDocPago.ProcesaPlantilla(txtPlantilla, False, lsMovNro, vdFechaDoc, lsEntiOrig, lsEntiDest, vnImporte, lsSimbolo, lsCtaEntiOrig, lsCtaEntiDest, vsNroDoc, gnColPage, lnMagIzq, lnMagDer, Mid(lsOpeCod, 3, 1), lnMagSup)
   If pageDoc.Tab = 0 Then txtNroCarta.SetFocus
   If pageDoc.Tab = 1 Then txtNroCH.SetFocus
   If pageDoc.Tab = 2 Then txtNroOP.SetFocus
End If
End Sub

Private Sub txtDocFecha_Validate(Cancel As Boolean)
If ValFecha(txtDocFecha) = False Then
    Cancel = True
End If
vdFechaDoc = txtDocFecha
rtfDoc.Text = oDocPago.ProcesaPlantilla(txtPlantilla, False, lsMovNro, vdFechaDoc, lsEntiOrig, lsEntiDest, vnImporte, lsSimbolo, lsCtaEntiOrig, lsCtaEntiDest, vsNroDoc, gnColPage, lnMagIzq, lnMagDer, Mid(lsOpeCod, 3, 1), lnMagSup)
End Sub

Private Sub txtGlosa_GotFocus()
txtGlosa.SelStart = 0
txtGlosa.SelLength = Len(txtGlosa.Text)
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Select Case pageDoc.Tab
        Case 0
            txtNroCarta.SetFocus
        Case 1
             txtNroCH.SetFocus
        Case 2
            If txtNroOP.Enabled Then
                txtNroOP.SetFocus
            Else
                If txtPersOP.Enabled Then
                    txtPersOP.SetFocus
                Else
                    cmdAceptar.Enabled = True
                    cmdAceptar.SetFocus
                End If
            End If
    End Select
   'If pageDoc.Tab = 0 Then txtNroCarta.SetFocus
   'If pageDoc.Tab = 1 Then txtNroCH.SetFocus
   'If pageDoc.Tab = 2 Then txtNroOP.SetFocus
End If
End Sub

Private Sub txtNroCarta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtNroCarta_Validate (False)
   If cmdAceptar.Enabled Then
       cmdAceptar.SetFocus
    End If
End If
End Sub

Private Sub txtNroCarta_Validate(Cancel As Boolean)
Dim oDoc As DDocumento
Set oDoc = New DDocumento
If Len(Trim(txtNroCarta.Text)) <> 0 Then
    If ValFecha(txtDocFecha) = False Then Exit Sub
   Cancel = False
   vsNroDoc = Format(txtNroCarta.Text, String(8, "0"))
   txtNroCarta.Text = vsNroDoc

    If oDoc.GetValidaDocProv("", lnDocTpo, txtNroCarta) Then
        MsgBox "Carta ya ha sido Ingresada", vbInformation, "Aviso"
        Cancel = True
        Exit Sub
    End If
   cmdAceptar.Enabled = True
   vdFechaDoc = txtDocFecha
   rtfDoc.Text = oDocPago.ProcesaPlantilla(txtPlantilla, False, lsMovNro, vdFechaDoc, lsEntiOrig, lsEntiDest, vnImporte, lsSimbolo, lsCtaEntiOrig, lsCtaEntiDest, vsNroDoc, gnColPage, lnMagIzq, lnMagDer, Mid(lsOpeCod, 3, 1), lnMagSup)

Else
   MsgBox "Número de carta no válido...", vbInformation, "Aviso"
   Cancel = True
End If
Set oDoc = Nothing
End Sub

Private Sub txtNroCH_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    vsNroDoc = txtNroCH.Text
    cmdAceptar.Enabled = True
    If lbIngresoPers Then
        txtPersCH.SetFocus
    Else
        cmdAceptar.SetFocus
    End If
End If
End Sub

'Private Sub txtNroCH_KeyUp(KeyCode As Integer, Shift As Integer)
'If Len(Trim(txtNroCH)) > 0 Then
'   cmdAceptar.Enabled = True
'End If
'End Sub

'Private Sub txtNroCH_Validate(Cancel As Boolean)
'If Len(Trim(txtNroCH.Text)) <> 0 Then
'   Cancel = False
'   vsNroDoc = txtNroCH.Text
'   cmdAceptar.Enabled = True
'   If lbIngresoPers Then
'      txtPersCH.SetFocus
'   Else
'      cmdAceptar.SetFocus
'   End If
'Else
'   MsgBox "Número de Cheque no válido...", vbInformation, "Aviso"
'   Cancel = True
'End If
'End Sub

Private Sub txtNroOP_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtNroOP_Validate (False)
End If
End Sub

Private Sub txtNroOP_KeyUp(KeyCode As Integer, Shift As Integer)
If Len(Trim(txtNroOP)) > 0 Then
   cmdAceptar.Enabled = True
End If
End Sub

Private Sub txtNroOP_LostFocus()
If Len(Trim(txtNroOP)) > 0 Then
    txtNroOP = Format(txtNroOP.Text, String(8, "0"))
End If
End Sub

Private Sub txtNroOP_Validate(Cancel As Boolean)
If Len(Trim(txtNroOP.Text)) <> 0 Then
   Cancel = False
   vsNroDoc = txtNroOP.Text
   cmdAceptar.Enabled = True
   If lbIngresoPers Then
      txtPersOP.SetFocus
   Else
      cmdAceptar.SetFocus
   End If
Else
   MsgBox "Número de Orden de Pago no válido...", vbCritical, "Error"
   Cancel = True
End If
End Sub

Private Sub txtPersCH_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   If txtCantCH.Enabled Then
      txtCantCH.SetFocus
   Else
      If cmdAceptar.Enabled Then cmdAceptar.SetFocus
   End If
End If
End Sub
Private Sub txtPersOP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   cmdAceptar.Enabled = True
   If txtCantOP.Enabled Then
      txtCantOP.SetFocus
   Else
      If cmdAceptar.Enabled Then
         cmdAceptar.SetFocus
      End If
   End If
End If
End Sub
Private Sub txtVoucherCh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtVoucherCh = txtVoucherCh
   If cmdAceptar.Enabled Then
      cmdAceptar.SetFocus
   Else
      cmdCancelar.SetFocus
   End If
Else
   If KeyAscii <> 8 Then
      KeyAscii = IIf(InStr("0123456789", Chr(KeyAscii)) > 0, KeyAscii, 0)
   End If
End If
End Sub

Private Sub txtVoucherOP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtVoucherOP = Format(Val(txtVoucherOP), "00000000")
Else
   If KeyAscii <> 8 Then
      KeyAscii = IIf(InStr("0123456789", Chr(KeyAscii)) > 0, KeyAscii, 0)
   End If
End If
End Sub
'ALPA 20090323****************************
Public Sub InicioPenalidad(ByVal psOpeCod As String, ByVal psOpeDesc As String, psPersNombre As String, psGlosa As String, pnImporte As Currency, pdFechaSist As Date, Optional psDocNroVoucher As String = "", Optional pbIngPers As Boolean, Optional psCodAge As String = "", Optional pbModificaMonto As Boolean = False)
Pag0 = False
Pag1 = False
Pag2 = True
lbIngresoPers = pbIngPers
lsOpeCod = psOpeCod
ldFechaSist = pdFechaSist
lsOpeDesc = psOpeDesc
vnImporte = pnImporte
vsGlosa = psGlosa
vsPersNombre = psPersNombre
vsNroVoucher = psDocNroVoucher
vsCodAge = psCodAge
lbModificaMonto = pbModificaMonto
vdFechaDoc = Format(ldFechaSist, "dd/mm/yyyy")
'vsTpoDoc = Format(TpoDocRetenciones, "00")
vbOk = True
vsNroVoucher = "0"
vsFormaDoc = ""
'Me.Show 1
End Sub
'*****************************************


