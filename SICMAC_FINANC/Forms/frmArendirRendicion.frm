VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5F774E03-DB36-4DFC-AAC4-D35DC9379F2F}#1.1#0"; "VertMenu.ocx"
Begin VB.Form frmArendirRendicion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rendicion a Caja General"
   ClientHeight    =   5340
   ClientLeft      =   210
   ClientTop       =   2070
   ClientWidth     =   10920
   Icon            =   "frmArendirRendicion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   585
      Left            =   1590
      TabIndex        =   30
      Top             =   4680
      Width           =   3195
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   1920
         TabIndex        =   31
         Top             =   180
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Rendición"
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
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1725
      End
   End
   Begin VB.CommandButton cmdRendir 
      Caption         =   "&Rendicion"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   19
      Top             =   4830
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9450
      TabIndex        =   18
      Top             =   4830
      Width           =   1305
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   720
      Left            =   1575
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   3960
      Width           =   9210
   End
   Begin VB.Frame Frame1 
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
      Left            =   1575
      TabIndex        =   10
      Top             =   765
      Width           =   9225
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Agencia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4740
         TabIndex        =   25
         Top             =   210
         Width           =   690
      End
      Begin VB.Label lblAgeCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   315
         Left            =   5475
         TabIndex        =   24
         Top             =   150
         Width           =   525
      End
      Begin VB.Label lblAgeDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   315
         Left            =   6030
         TabIndex        =   23
         Top             =   150
         Width           =   2940
      End
      Begin VB.Label lblAreaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   315
         Left            =   1485
         TabIndex        =   22
         Top             =   150
         Width           =   3180
      End
      Begin VB.Label lblAreaCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   315
         Left            =   855
         TabIndex        =   21
         Top             =   150
         Width           =   585
      End
      Begin VB.Label label5 
         AutoSize        =   -1  'True
         Caption         =   "Area : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   225
         TabIndex        =   20
         Top             =   165
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Persona"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   225
         TabIndex        =   13
         Top             =   555
         Width           =   600
      End
      Begin VB.Label txtPerCod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   300
         Left            =   855
         TabIndex        =   12
         Top             =   510
         Width           =   1545
      End
      Begin VB.Label txtPerNom 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   300
         Left            =   2460
         TabIndex        =   11
         Top             =   510
         Width           =   5415
      End
   End
   Begin VB.Frame fraDocEmit 
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
      Left            =   3435
      TabIndex        =   3
      Top             =   60
      Width           =   7335
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5970
         TabIndex        =   15
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Saldo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5415
         TabIndex        =   14
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   105
         TabIndex        =   9
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2010
         TabIndex        =   8
         Top             =   315
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3570
         TabIndex        =   7
         Top             =   300
         Width           =   435
      End
      Begin VB.Label txtDocNro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   420
         TabIndex        =   6
         Top             =   240
         Width           =   1530
      End
      Begin VB.Label txtRecImporte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4050
         TabIndex        =   5
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label txtDocFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "dd/mm/yyyy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2505
         TabIndex        =   4
         Top             =   255
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Recibo de A rendir"
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
      Left            =   1575
      TabIndex        =   0
      Top             =   60
      Width           =   2070
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   2
         Top             =   285
         Width           =   255
      End
      Begin VB.Label txtRecNro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   450
         TabIndex        =   1
         Top             =   225
         Width           =   1380
      End
   End
   Begin VB.Frame FraCtaIFPagadora 
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
      ForeColor       =   &H000040C0&
      Height          =   2220
      Left            =   1575
      TabIndex        =   26
      Top             =   1680
      Width           =   9240
      Begin Sicmact.TxtBuscar txtBuscaEntidad 
         Height          =   375
         Left            =   435
         TabIndex        =   27
         Top             =   375
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   661
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCtaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   435
         TabIndex        =   29
         Top             =   1185
         Width           =   8115
      End
      Begin VB.Label lblIFNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   435
         TabIndex        =   28
         Top             =   810
         Width           =   8115
      End
   End
   Begin VB.Frame fraExacta 
      Caption         =   "Rendicion Exacta"
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
      Height          =   2220
      Left            =   1575
      TabIndex        =   16
      Top             =   1680
      Width           =   9240
   End
   Begin VertMenu.VerticalMenu VMTpoRendAge 
      Height          =   5100
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   8996
      MenusMax        =   3
      MenuCaption1    =   "Exacta"
      MenuItemIcon11  =   "frmArendirRendicion.frx":030A
      MenuItemCaption11=   "Exacta"
      MenuCaption2    =   "Exacta"
      MenuItemIcon21  =   "frmArendirRendicion.frx":0624
      MenuItemCaption21=   "Exacta"
      MenuCaption3    =   "Egresos"
      MenuItemsMax3   =   3
      MenuItemIcon31  =   "frmArendirRendicion.frx":093E
      MenuItemCaption31=   "Ventanilla"
      MenuItemIcon32  =   "frmArendirRendicion.frx":0C58
      MenuItemCaption32=   "Ajuste de Saldo"
      MenuItemIcon33  =   "frmArendirRendicion.frx":0F72
      MenuItemCaption33=   "Item3"
   End
   Begin VertMenu.VerticalMenu VMTpoRend 
      Height          =   5100
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   8996
      MenusMax        =   3
      MenuCaption1    =   "Exacta"
      MenuItemIcon11  =   "frmArendirRendicion.frx":128C
      MenuItemCaption11=   "Exacta"
      MenuCaption2    =   "Egresos"
      MenuItemsMax2   =   6
      MenuItemIcon21  =   "frmArendirRendicion.frx":15A6
      MenuItemCaption21=   "Efectivo"
      MenuItemIcon22  =   "frmArendirRendicion.frx":18C0
      MenuItemCaption22=   "Cheque"
      MenuItemIcon23  =   "frmArendirRendicion.frx":1BDA
      MenuItemCaption23=   "Orden Pago"
      MenuItemIcon24  =   "frmArendirRendicion.frx":1EF4
      MenuItemCaption24=   "Carta"
      MenuItemIcon25  =   "frmArendirRendicion.frx":220E
      MenuItemCaption25=   "Nota de Abono"
      MenuItemIcon26  =   "frmArendirRendicion.frx":2528
      MenuItemCaption26=   "Otros Egresos"
      MenuCaption3    =   "Ingresos"
      MenuItemsMax3   =   6
      MenuItemIcon31  =   "frmArendirRendicion.frx":2842
      MenuItemCaption31=   "Efectivo"
      MenuItemIcon32  =   "frmArendirRendicion.frx":2B5C
      MenuItemCaption32=   "Cheque"
      MenuItemIcon33  =   "frmArendirRendicion.frx":2E76
      MenuItemCaption33=   "Orden Pago"
      MenuItemIcon34  =   "frmArendirRendicion.frx":3190
      MenuItemCaption34=   "por Ventanilla"
      MenuItemIcon35  =   "frmArendirRendicion.frx":34AA
      MenuItemCaption35=   "Nota Cargo"
      MenuItemIcon36  =   "frmArendirRendicion.frx":37C4
      MenuItemCaption36=   "Otros Ingresos"
   End
End
Attribute VB_Name = "frmArendirRendicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsNroArendir As String
Dim lsNroDoc As String
Dim lsFechaDoc As String
Dim lsPersCod As String
Dim lsPersNomb As String
Dim lsAreaCod As String
Dim lsAreaDesc As String
Dim lsDescDoc As String
Dim lnImporte As Currency
Dim lsAgeCod As String
Dim lsAgeDesc As String
Public lnSaldo As Currency
Dim lsMovNroAtencion As String
Dim lsMovNroSol As String
Dim lsCtaContArendir As String
Dim lsCtaContPendiente As String
Dim lnFaseArendir As ARendirFases
Dim lnTipoArendir As ArendirTipo
Dim lnMenuMumber  As Long
Dim lnMenuItem As Long
Dim lsDocTpo As TpoDoc
Dim lsGlosa As String
Dim lsAbreDocArendir As String

Dim oAreas As DActualizaDatosArea
Dim oArendir As NARendir
Dim oCtasIF As NCajaCtaIF
Dim oOpe As DOperacion


Dim lbLoad As Boolean
Dim lbOk As Boolean

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************



Public Sub Inicio(ByVal pnFaseArendir As ARendirFases, ByVal pnTipoArendir As ArendirTipo, _
                ByVal psNroArendir As String, ByVal psNroDoc As String, ByVal psFechaDoc As String, _
                ByVal psPersCod As String, ByVal psPersNomb As String, ByVal psAreaCod As String, _
                ByVal psAreaDesc As String, ByVal psAgeCod As String, ByVal psAgeDesc As String, ByVal psDescDoc As String, _
                ByVal psMovNroAtencion As String, ByVal psAbreDocArendir As String, _
                ByVal pnImporte As Currency, ByVal psCtaContARendir As String, ByVal psCtaContPendiente As String, _
                ByVal pnSaldo As Currency, ByVal psMovNroSol As String, ByVal psGlosa As String)


lnFaseArendir = pnFaseArendir
lsNroArendir = psNroArendir
lsAbreDocArendir = psAbreDocArendir
lsNroDoc = psNroDoc
lsFechaDoc = psFechaDoc
lsPersCod = psPersCod
lsPersNomb = psPersNomb
lsAreaCod = psAreaCod
lsGlosa = psGlosa
lsAreaDesc = psAreaDesc
lsAgeCod = psAgeCod
lsAgeDesc = psAgeDesc
lsDescDoc = psDescDoc
lnImporte = pnImporte
lsCtaContPendiente = psCtaContPendiente
lsCtaContArendir = psCtaContARendir
lsMovNroAtencion = psMovNroAtencion
lsMovNroSol = psMovNroSol
lnSaldo = pnSaldo
lnTipoArendir = pnTipoArendir

Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
lbOk = False
Unload Me
End Sub
Private Sub cmdRendir_Click()
Dim rs As ADODB.Recordset
Dim oContFunc As NContFunciones
Dim oDocPago As clsDocPago
Dim oDocRec As NDocRec
Dim oContImp As NContImprimir
                

Dim lsMovNro As String
Dim lsOpeCod As String
Dim lbMueveCtasCont As Boolean
Dim lsCtaPendiente As String
Dim lsCtaOperacion As String
Dim lsCtaDiferencia As String
Dim lnImporte  As Currency
Dim lnMontoDif As Currency
Dim lsDocVoucher As String
Dim lsDocumento As String
Dim lbEfectivo As Boolean
Dim lbIngreso As Boolean

Dim lsPersCodIf As String
Dim lsTipoIF As String
Dim lsCtaBanco As String

Dim lsCtaChqIf As String
Dim lnPlaza As Integer
Dim ldValchq As Date
Dim lsDocNroVoucher As String
Dim ldFechaVoucher As Date
Dim lsDocNro As String
Dim lsEntidadOrig As String
Dim lsCtaEntidadOrig As String
Dim lsGlosa As String
Dim lsPersNombre As String
Dim lsSubCuentaIF As String
Dim lsPersCod As String
Dim lsFechaDoc As String
Dim lsDocViaticos As String
Dim lsCuentaAho As String
Dim lnPersoneria As PersPersoneria

Dim lsPersDireccion As String
Dim lsUbigeo As String
Dim lnMotivo As MotivoNotaAbonoCargo
Dim lsCadBol As String

Dim lsAgeDevolucion As String

On Error GoTo ErrGraba
lsEntidadOrig = lblIFNombre
lsCtaEntidadOrig = Trim(lblCtaDesc)
lsPersNombre = txtPerNom
lsPersCod = txtPerCod
lsSubCuentaIF = oCtasIF.SubCuentaIF(Mid(txtBuscaEntidad.Text, 1, 13))
lnImporte = CCur(lblSaldo)
lsDocNro = ""
lsDocVoucher = ""
If lnTipoArendir = gArendirTipoViaticos Then
    lsDocViaticos = lsAbreDocArendir + "-" + lsNroArendir
Else
    lsDocViaticos = ""
End If

lsCtaBanco = ""
lsPersCodIf = ""
lsTipoIF = ""
Set oDocPago = New clsDocPago
Set oOpe = New DOperacion
'Dim oCap As NCapMovimientos
'Set oCap = New NCapMovimientos

lbMueveCtasCont = True
If ValidaDatos = False Then Exit Sub
Set oContFunc = New NContFunciones
cmdRendir.Enabled = False
lsDocTpo = -1
lbIngreso = False
lsCtaDiferencia = ""
lsCuentaAho = ""
lnMontoDif = 0
Set oDocRec = New NDocRec
Set oContImp = New NContImprimir
Select Case lnMenuMumber
    Case 1   'Exacta
        lbMueveCtasCont = False

    Case 2   'Egresos
        If txtFecha <> gdFecSis Then
            MsgBox "Fecha de Rendición tiene que ser " & gdFecSis, vbInformation, "¡Aviso!"
            Exit Sub
        End If
        Select Case lnMenuItem
            Case 1    'Efectivo
                frmArendirEfectivo.Inicio lnTipoArendir, lsNroArendir, Mid(gsOpeCod, 3, 1), lsAreaDesc, lnSaldo, lsPersCod, lsPersNomb, lnFaseArendir
                If Not frmArendirEfectivo.lbOk Then
                    Unload frmArendirEfectivo
                    Set frmArendirEfectivo = Nothing
                    Exit Sub
                End If
                Set rs = frmArendirEfectivo.rsEfectivo
                If frmArendirEfectivo.vnDiferencia <> 0 Then
                    lnMontoDif = frmArendirEfectivo.vnDiferencia
                End If
                Unload frmArendirEfectivo
                Set frmArendirEfectivo = Nothing
            Case 2    'Cheque
                Screen.MousePointer = 11
                lsDocTpo = TpoDocCheque
                lsCtaBanco = Mid(txtBuscaEntidad, 18, Len(txtBuscaEntidad))
                lsPersCodIf = Mid(txtBuscaEntidad, 4, 13)
                lsTipoIF = Mid(txtBuscaEntidad, 1, 2)
                'oDocPago.InicioCheque lsDocNro, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeDesc, txtMovDesc, lnImporte, gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocNroVoucher, False, gsCodAge, lsCtaBanco
                oDocPago.InicioCheque lsDocNro, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeDesc, txtMovDesc, lnImporte, gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocNroVoucher, False, gsCodAge, lsCtaBanco, , lsTipoIF, lsPersCodIf, lsCtaBanco
                Screen.MousePointer = 0
                If oDocPago.vbOk Then    'Se ingresó dato de Cheque u Orden de Pago
                    lsFechaDoc = oDocPago.vdFechaDoc
                    lsDocNro = oDocPago.vsNroDoc
                    lsDocNroVoucher = oDocPago.vsNroVoucher
                    ldFechaVoucher = oDocPago.vdFechaDoc
                    lsDocumento = oDocPago.vsFormaDoc
                    txtMovDesc = oDocPago.vsGlosa
                 Else
                    Exit Sub
                End If
            Case 3    'Orden Pago
                Screen.MousePointer = 11
                lsDocTpo = TpoDocOrdenPago
                oDocPago.InicioOrdenPago lsDocNro, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeCod, txtMovDesc, lnImporte, gdFecSis, "", False, gsCodAge
                Screen.MousePointer = 0
                If oDocPago.vbOk Then    'Se ingresó dato de Cheque u Orden de Pago
                    lsFechaDoc = oDocPago.vdFechaDoc
                    lsDocNro = oDocPago.vsNroDoc
                    lsDocNroVoucher = oDocPago.vsNroVoucher
                    ldFechaVoucher = oDocPago.vdFechaDoc
                    lsDocumento = oDocPago.vsFormaDoc
                    txtMovDesc = oDocPago.vsGlosa
                Else
                    Exit Sub
                End If
            Case 4    'Carta
                lsDocTpo = TpoDocCarta
                lsCtaBanco = Mid(txtBuscaEntidad, 18, Len(txtBuscaEntidad))
                lsPersCodIf = Mid(txtBuscaEntidad, 4, 13)
                lsTipoIF = Mid(txtBuscaEntidad, 1, 2)
                oDocPago.InicioCarta lsDocNro, lsPersCod, gsOpeCod, gsOpeCod, txtMovDesc, "", lnImporte, gdFecSis, lsEntidadOrig, lsCtaEntidadOrig, lsPersNombre, "", ""
                If oDocPago.vbOk Then    'Se ingresó datos de carta
                    lsFechaDoc = oDocPago.vdFechaDoc
                    lsDocNro = oDocPago.vsNroDoc
                    lsDocNroVoucher = oDocPago.vsNroVoucher
                    ldFechaVoucher = oDocPago.vdFechaDoc
                    lsDocumento = oDocPago.vsFormaDoc
                    txtMovDesc = oDocPago.vsGlosa
                Else
                    Exit Sub
                End If
            Case 5    'Nota de Abono
                Dim oImp As New NContImprimir
                Dim oDis As New NRHProcesosCierre
                lsDocTpo = TpoDocNotaAbono
                frmNotaCargoAbono.Inicio TpoDocNotaAbono, CCur(lblSaldo), gdFecSis, txtMovDesc, gsOpeCod
                If frmNotaCargoAbono.vbOk Then
                    lsDocNro = frmNotaCargoAbono.NroNotaCA
                    lsFechaDoc = frmNotaCargoAbono.FechaNotaCA
                    txtMovDesc = frmNotaCargoAbono.Glosa
                    lsDocumento = frmNotaCargoAbono.NotaCargoAbono
                    lsPersNombre = frmNotaCargoAbono.lblpersNombre
                    lsPersDireccion = frmNotaCargoAbono.lblPersDireccion
                    lsUbigeo = frmNotaCargoAbono.lblUbigeo
                    ldFechaVoucher = frmNotaCargoAbono.FechaNotaCA
                    lsCuentaAho = frmNotaCargoAbono.CuentaAhoNro

                    lsDocumento = oImp.ImprimeNotaAbono(Format(ldFechaVoucher, gsFormatoFecha), lnImporte, txtMovDesc, lsCuentaAho, lsPersNombre)
                    lsCadBol = oDis.ImprimeBoletaCad(ldFechaVoucher, "ABONO CAJA GENERAL", "Depósito CAJA GENERAL*Nro." & lsDocNro, "", lnImporte, lsPersNombre, lsCuentaAho, "", 0, 0, "Nota Abono", 0, 0, False, False, , , , True, , , , False, gsNomAge) & oImpresora.gPrnSaltoPagina

                    Unload frmNotaCargoAbono
                    Set frmNotaCargoAbono = Nothing
                Else
                    Unload frmNotaCargoAbono
                    Set frmNotaCargoAbono = Nothing
                    Exit Sub
                End If
            Case 6   'Otros Egresos
                'validacion dentro del menu
        End Select
    Case 3   'Ingresos
        lbIngreso = True
        If Not lnMenuItem = 4 Then
            If txtFecha <> gdFecSis Then
                MsgBox "Fecha de Rendición tiene que ser " & gdFecSis, vbInformation, "¡Aviso!"
                Exit Sub
            End If
        End If
        Select Case lnMenuItem
            Case 1    'Efectivo
                frmArendirEfectivo.Inicio lnTipoArendir, lsNroArendir, Mid(gsOpeCod, 3, 1), lsAreaDesc, lnSaldo, lsPersCod, lsPersNomb, lnFaseArendir
                If Not frmArendirEfectivo.lbOk Then
                    Exit Sub
                End If
                Set rs = frmArendirEfectivo.rsEfectivo
                If frmArendirEfectivo.vnDiferencia <> 0 Then
                    lnMontoDif = frmArendirEfectivo.vnDiferencia
                End If
            Case 2    'Cheque
                'Registro de cheque
                'EJVG20140415 ***
                'Set frmIngCheques = Nothing
                'lsDocTpo = TpoDocCheque
                'lsOpeCod = oArendir.GetOpeRendicion(Mid(gsOpeCod, 1, 5), lsDocTpo, lsCtaContArendir, lsCtaContPendiente, lbMueveCtasCont, IIf(lbIngreso = True, "D", "H"))
                'lsCtaPendiente = oOpe.EmiteOpeCta(lsOpeCod, "H")
                'frmIngCheques.InicioArendir lsOpeCod, lnImporte, lnTipoArendir, lsMovNroAtencion, lsMovNroSol, lsCtaPendiente, Trim(txtMovDesc), Mid(gsOpeCod, 3, 1)
                'If frmIngCheques.OK = False Then
                '    Exit Sub
                'Else
                '    lbOk = True
                '    Unload Me
                '    Exit Sub
                'End If
                Exit Sub
                'END EJVG *******
            Case 3    'Orden Pago
                lsDocTpo = TpoDocOrdenPago
                frmNotaCargoAbono.Inicio TpoDocOrdenPago, CCur(lblSaldo), gdFecSis, txtMovDesc, gsOpeCod
                If frmNotaCargoAbono.vbOk Then
                    lsDocNro = frmNotaCargoAbono.NroNotaCA
                    lsFechaDoc = frmNotaCargoAbono.FechaNotaCA
                    txtMovDesc = frmNotaCargoAbono.Glosa
                    lsDocumento = frmNotaCargoAbono.NotaCargoAbono
                    lsPersNombre = frmNotaCargoAbono.PersNombre
                    lsPersDireccion = frmNotaCargoAbono.PersDireccion
                    lsUbigeo = frmNotaCargoAbono.PersUbigeo
                    lnPersoneria = frmNotaCargoAbono.Personeria
                    ldFechaVoucher = frmNotaCargoAbono.FechaNotaCA
                    lsCuentaAho = frmNotaCargoAbono.CuentaAhoNro
                Else
                    Exit Sub
                End If
            
            Case 4    'Ingresos por Ventanilla
                lsDocTpo = TpoDocRecibosDiversos
                frmOpeNegVentanilla.Inicio lsNroArendir, Mid(gsOpeCod, 3, 1), lnSaldo, lsPersCod, lsPersNomb
                If Not frmOpeNegVentanilla.lbOk Then
                    Exit Sub
                End If
                
                lsAgeDevolucion = frmOpeNegVentanilla.lsAgeCodRef
                Set rs = frmOpeNegVentanilla.rsPago
                If frmOpeNegVentanilla.vnDiferencia <> 0 Then
                    lnMontoDif = frmOpeNegVentanilla.vnDiferencia
                End If
                
            Case 5    'Nota de Cargo
                lsDocTpo = TpoDocNotaCargo
                frmNotaCargoAbono.Inicio TpoDocNotaCargo, CCur(lblSaldo), gdFecSis, txtMovDesc, gsOpeCod
                If frmNotaCargoAbono.vbOk Then
                    lsDocNro = frmNotaCargoAbono.NroNotaCA
                    lsFechaDoc = frmNotaCargoAbono.FechaNotaCA
                    txtMovDesc = frmNotaCargoAbono.Glosa
                    lsDocumento = frmNotaCargoAbono.NotaCargoAbono
                    'lsDocNroVoucher = oContFunc.GeneraDocNro(TpoDocVoucherEgreso, Mid(gsOpeCod, 3, 1))
                    lsPersNombre = frmNotaCargoAbono.PersNombre
                    lsPersDireccion = frmNotaCargoAbono.PersDireccion
                    lsUbigeo = frmNotaCargoAbono.PersUbigeo
                    ldFechaVoucher = frmNotaCargoAbono.FechaNotaCA
                    lsCuentaAho = frmNotaCargoAbono.CuentaAhoNro
                Else
                    Exit Sub
                End If
            Case 6    'Otros Ingresos
                'validacion dentro del menu
        End Select
End Select
'***Modificado por ELRO el 20120508, según OYP-RFC005-2012 y OYP-RFC016-2012
'lsOpeCod = oArendir.GetOpeRendicion(Mid(gsOpeCod, 1, 4) & IIf(lbIngreso, "5", "6"), lsDocTpo, lsCtaContARendir, lsCtaContPendiente, lbMueveCtasCont, IIf(lbIngreso = True, "D", "H"))
lsOpeCod = oArendir.GetOpeRendicion(Mid(gsOpeCod, 1, 4) & IIf(gsOpeCod = CStr(gCGArendirViatRendMN) Or gsOpeCod = CStr(gCGArendirViatRendME) Or gsOpeCod = CStr(gCGArendirCtaRendMN) Or gsOpeCod = CStr(gCGArendirCtaRendME), "5", "6"), lsDocTpo, lsCtaContArendir, lsCtaContPendiente, lbMueveCtasCont, IIf(lbIngreso = True, "D", "H"))
'***Fin Modificado por ELRO*******************************

'*** PEAC 20440622
Dim oCont As New NContFunciones
If Not oCont.PermiteModificarAsiento(Format(Me.txtFecha, gsFormatoMovFecha), False) Then
   MsgBox "No se puede procesar con fecha de un Mes Contable Cerrado.", vbInformation, "Atención"
   Exit Sub
End If
Set oCont = Nothing
'*** FIN PEAC

If MsgBox("Desea Grabar la Rendicion respectiva?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oContFunc.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
    lbEfectivo = False
    Select Case lnMenuMumber
        Case 1  'Exacta
            lsOpeCod = oArendir.GetOpeRendicion(Mid(gsOpeCod, 1, 5), lsDocTpo, lsCtaContArendir, lsCtaContPendiente, lbMueveCtasCont)
            oArendir.GrabaRendicionExacta lnTipoArendir, gsFormatoFecha, lsMovNro, lsOpeCod, txtMovDesc, lsMovNroAtencion, lsMovNroSol
'***Comentado por ELRO el 20120426, según OYP-RFC005-2012 y OYP-RFC016-2012
'        Case 2  'Egresos
'            lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "H")
'            lsCtaPendiente = oOpe.EmiteOpeCta(lsOpeCod, "D")
'
'            Select Case lnMenuItem
'                Case 1    'Efectivo
'                    lbEfectivo = True
'                    lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "H", "1")
'                    lsCtaPendiente = oOpe.EmiteOpeCta(lsOpeCod, "D", "1")
'                    lsCtaDiferencia = oOpe.EmiteOpeCta(lsOpeCod, IIf(lnMontoDif > 0, "D", "H"), "2")
'                    oArendir.GrabaRendicionEfectivo lnTipoArendir, gsFormatoFecha, lsMovNro, lsOpeCod, txtMovDesc, _
'                                                   lsCtaPendiente, lsCtaOperacion, lnImporte, rs, lsMovNroAtencion, lsMovNroSol, _
'                                                   lbIngreso, lsCtaDiferencia, lnMontoDif
'                Case 2    'Cheque
'                    lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "H", , txtBuscaEntidad, CtaOBjFiltroIF, True)
'
'                    oArendir.GrabaRendicionGiroDocumento lnTipoArendir, lsMovNro, lsMovNroSol, _
'                                    lsMovNroAtencion, lsOpeCod, txtMovDesc, lsCtaPendiente, _
'                                    lsCtaOperacion, lsPersCod, lnImporte, lsDocTpo, lsDocNro, Format(CDate(lsFechaDoc), gsFormatoFecha), lsDocNroVoucher, lsPersCodIf, lsTipoIF, _
'                                    lsCtaBanco
'
'                Case 3      'Orden Pago
'                    lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "H", , frmARendirLista.txtBuscarArendir, ObjCMACAgenciaArea)
'                    oArendir.GrabaRendicionGiroDocumento lnTipoArendir, lsMovNro, lsMovNroSol, _
'                                    lsMovNroAtencion, lsOpeCod, txtMovDesc, lsCtaPendiente, _
'                                    lsCtaOperacion, lsPersCod, lnImporte, lsDocTpo, lsDocNro, Format(CDate(lsFechaDoc), gsFormatoFecha), lsDocNroVoucher, lsPersCodIf, lsTipoIF, _
'                                    lsCtaBanco
'                Case 4      'Carta
'                    lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "H", "0", txtBuscaEntidad, CtaOBjFiltroIF, True)
'                    oArendir.GrabaRendicionGiroDocumento lnTipoArendir, lsMovNro, lsMovNroSol, _
'                                    lsMovNroAtencion, lsOpeCod, txtMovDesc, lsCtaPendiente, _
'                                    lsCtaOperacion, lsPersCod, lnImporte, lsDocTpo, lsDocNro, lsFechaDoc, lsDocNroVoucher, lsPersCodIf, lsTipoIF, _
'                                    lsCtaBanco
'                Case 5      'Nota de Abono
'                    lnMotivo = gNARendirCuenta
'                    'lsDocNro = oDocRec.GetNroNotaCargoAbono(TpoDocNotaAbono)
'                    oArendir.GrabaRendicionGiroDocumento lnTipoArendir, lsMovNro, lsMovNroSol, _
'                            lsMovNroAtencion, lsOpeCod, txtMovDesc, lsCtaPendiente, _
'                            lsCtaOperacion, lsPersCod, lnImporte, lsDocTpo, lsDocNro, Format(CDate(lsFechaDoc), gsFormatoFecha), lsDocNroVoucher, lsPersCodIf, lsTipoIF, _
'                            lsCtaBanco, , , , lnMotivo, lsCuentaAho, gbBitCentral, True
'
'
'                Case 6      'Otros Egresos
'                    lsOpeCod = gsOpeCod
'                    lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "H", "1")
'                    lsCtaPendiente = oOpe.EmiteOpeCta(lsOpeCod, "D", "1")
'                    oArendir.GrabaRendicionGiroDocumento lnTipoArendir, lsMovNro, lsMovNroSol, _
'                            lsMovNroAtencion, lsOpeCod, txtMovDesc, lsCtaPendiente, _
'                            lsCtaOperacion, lsPersCod, lnImporte, lsDocTpo, "", "", "", "", "", ""
'
'            End Select
'***Fin Comentado por ELRO*******************************
        Case 3 'Ingresos
            lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "D") '& lsAgeDevolucion
            lsCtaPendiente = oOpe.EmiteOpeCta(lsOpeCod, "H")
            Select Case lnMenuItem
                Case 1    'Efectivo
                    lsCtaDiferencia = oOpe.EmiteOpeCta(lsOpeCod, IIf(lnMontoDif > 0, "H", "D"), "2")
                    oArendir.GrabaRendicionEfectivo lnTipoArendir, gsFormatoFecha, lsMovNro, lsOpeCod, txtMovDesc, _
                                                lsCtaPendiente, lsCtaOperacion, lnImporte, rs, lsMovNroAtencion, lsMovNroSol, lbIngreso, lsCtaDiferencia, lnMontoDif
                    lbEfectivo = True
                Case 2    'Cheque
                    'Se realiza la grabación dentro del formulario de resgistro de cheques
                Case 3    'Orden Pago
                    lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "D", Trim(Str(lnPersoneria)), gsCodAge, ObjCMACAgencias)
                    lsCtaPendiente = oOpe.EmiteOpeCta(lsOpeCod, "H")
                    oArendir.CapCargoCuentaAhoMov gsFormatoFecha, lsCuentaAho, lnImporte, lsOpeCod, lsMovNro, txtMovDesc, TpoDocOrdenPago, _
                            lsDocNro, , True, , , , True, lnTipoArendir, lsCtaPendiente, lsCtaOperacion, lsMovNroAtencion, _
                            lsMovNroSol, gdFecSis, True
                            
                Case 4    'Ingreso por Ventanilla
                    lsCtaDiferencia = oOpe.EmiteOpeCta(lsOpeCod, IIf(lnMontoDif > 0, "H", "D"), "2")
                     oArendir.GrabaRendicionVentanilla lnTipoArendir, gsFormatoFecha, lsMovNro, lsOpeCod, txtMovDesc, _
                                                lsCtaPendiente, lsCtaOperacion, lnImporte, rs, lsMovNroAtencion, lsMovNroSol, lbIngreso, lsCtaDiferencia, lnMontoDif
                    lbEfectivo = False
                    
                Case 5    'Nota de Cargo
                    lsDocNro = oDocRec.GetNroNotaCargoAbono(TpoDocNotaCargo)
                    lnMotivo = gNCRendirCuenta
                    oArendir.GrabaRendicionGiroDocumento lnTipoArendir, lsMovNro, lsMovNroSol, _
                            lsMovNroAtencion, lsOpeCod, txtMovDesc, lsCtaOperacion, lsCtaPendiente, _
                            lsPersCod, lnImporte, lsDocTpo, lsDocNro, lsFechaDoc, lsDocNroVoucher, lsPersCodIf, lsTipoIF, _
                            lsCtaBanco, , , , lnMotivo, lsCuentaAho
                    
                    lsDocumento = oContImp.ImprimeNotaCargoAbono(lsDocNro, txtMovDesc, CCur(lnImporte), _
                                                lsPersNombre, lsPersDireccion, lsUbigeo, gdFecSis, Mid(gsOpeCod, 3, 1), lsCuentaAho, TpoDocNotaCargo, gsNomAge, gsCodUser)
                    
                Case 6    'Otros Ingresos
                    lsOpeCod = gsOpeCod
                    lsCtaOperacion = oOpe.EmiteOpeCta(lsOpeCod, "D", "0")
                    lsCtaPendiente = oOpe.EmiteOpeCta(lsOpeCod, "H", "0")
                    'edpyme
                    'oArendir.GrabaRendicionGiroDocumento lnTipoArendir, gsFormatoFecha, lsMovNro, lsMovNroSol, _
                    '         lsMovNroAtencion, lsOpeCod, txtMovDesc, lsCtaOperacion, _
                    '         lsCtaPendiente, lsPersCod, lnImporte, lsDocTpo, "", "", "", "", "", ""
                    oArendir.GrabaRendicionGiroDocumento lnTipoArendir, lsMovNro, lsMovNroSol, _
                             lsMovNroAtencion, lsOpeCod, txtMovDesc, lsCtaOperacion, _
                             lsCtaPendiente, lsPersCod, lnImporte, lsDocTpo, "", gsFormatoFecha, "", "", "", ""
                
                
            End Select
    End Select
    
    
    If lnMenuMumber <> 1 Then
        ImprimeAsientoContable lsMovNro, lsDocVoucher, lsDocTpo, lsDocumento, lbEfectivo, lbIngreso, txtMovDesc, lsPersCod, lnImporte, lnTipoArendir, lsDocViaticos, , , , "17", , , lsCadBol
    End If
    Set frmArendirEfectivo = Nothing
    lbOk = True
   ' Set oCap = Nothing
    Set oDocRec = Nothing
    Set oContImp = Nothing
    cmdRendir.Enabled = True
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, Me.Caption & " Recibo a Rendir : " & txtRecNro & txtDocNro & " el Monto : " & txtRecImporte
            Set objPista = Nothing
            '*******
    Unload Me
End If
Exit Sub
ErrGraba:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub
Private Sub Form_Load()
lblAgeCod = lsAgeCod
lblAgeDesc = lsAgeDesc
lblAreaCod = lsAreaCod
lblAreaDesc = lsAreaDesc
txtPerCod = lsPersCod
txtPerNom = lsPersNomb
txtDocNro = lsNroDoc
txtRecNro = lsNroArendir
txtDocFecha = lsFechaDoc
txtRecImporte = Format(lnImporte, "#,#0.00")
lblSaldo = Format(Abs(lnSaldo), "#,#0.00")
Set oCtasIF = New NCajaCtaIF
Dim oOpe As DOperacion

Set oOpe = New DOperacion
txtBuscaEntidad.psRaiz = "Cuentas de Instituciones Financieras"
'txtBuscaEntidad.rs = oCtasIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), gTpoIFBanco, gTpoCtaIFCtaCte & "-" & gTpoCtaIFCtaAho)
txtBuscaEntidad.rs = oOpe.GetRsOpeObj(gsOpeCod, "1")

txtMovDesc = lsGlosa
txtFecha = gdFecSis
lbLoad = True
CentraForm Me
If Val(lnSaldo) < 0 Then
    lblSaldo.ForeColor = &HFF&
Else
    lblSaldo.ForeColor = &HC00000
End If
If lnTipoArendir = gArendirTipoAgencias Then
    Select Case CCur(lnSaldo)
        Case Is = 0
            VMTpoRendAge.MenusMax = 1
            VMTpoRendAge.MenuCur = 1
            lnMenuMumber = 1
            lnMenuItem = 1
            VMTpoRendAge_MenuItemClick 1, 1
         Case Is < 0
            VMTpoRendAge.MenusMax = 2
            VMTpoRendAge.MenuCur = 2
            lnMenuMumber = 2
            lnMenuItem = 1
            VMTpoRendAge_MenuItemClick 2, 1
        Case Is > 0
            VMTpoRendAge.MenusMax = 3
            VMTpoRendAge.MenuCur = 3
            lnMenuMumber = 3
            lnMenuItem = 1
            VMTpoRendAge_MenuItemClick 3, 1
    End Select
Else
    Select Case CCur(lnSaldo)
        Case Is = 0
            VMTpoRend.MenusMax = 1
            VMTpoRend.MenuCur = 1
            lnMenuMumber = 1
            lnMenuItem = 1
            VMTpoRend_MenuItemClick 1, 1
        '***Comentado por ELRO el 20120425, según OYP-RFC005-2012 y OYP-RFC016-2012
        'Case Is < 0
        '    VMTpoRend.MenusMax = 2
        '    VMTpoRend.MenuCur = 2
        '    lnMenuMumber = 2
        '    lnMenuItem = 1
        '    VMTpoRend_MenuItemClick 2, 1
        '***Fin Comentado por ELRO*******************************
        Case Is > 0
            VMTpoRend.MenusMax = 3
            VMTpoRend.MenuCur = 3
            lnMenuMumber = 3
            lnMenuItem = 1
            VMTpoRend_MenuItemClick 3, 1
    End Select
End If
lbLoad = False

Me.VMTpoRendAge.Visible = False
Me.VMTpoRend.Visible = False
If lnTipoArendir = gArendirTipoAgencias Then
    Me.VMTpoRendAge.Visible = True
Else
    Me.VMTpoRend.Visible = True
End If
Set oAreas = New DActualizaDatosArea
Set oArendir = New NARendir

End Sub
Private Sub Form_Unload(Cancel As Integer)
Set oAreas = Nothing
Set oArendir = Nothing
Set oCtasIF = Nothing
End Sub

Private Sub txtBuscaEntidad_EmiteDatos()
If txtBuscaEntidad.Text <> "" Then
    lblIFNombre = oCtasIF.NombreIF(Mid(txtBuscaEntidad.Text, 4, 13))
    lblCtaDesc = oCtasIF.EmiteTipoCuentaIF(Mid(txtBuscaEntidad.Text, 18, Len(txtBuscaEntidad.Text))) & " " & txtBuscaEntidad.psDescripcion
    
    lblIFNombre = oCtasIF.NombreIF(Mid(txtBuscaEntidad.Text, 4, 13))
    lblCtaDesc = oCtasIF.EmiteTipoCuentaIF(Mid(txtBuscaEntidad.Text, 18, Len(txtBuscaEntidad.Text))) & " " & txtBuscaEntidad.psDescripcion
    cmdRendir_Click
Else
    lblIFNombre = ""
    lblCtaDesc = ""
End If
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValidaFechaContab(txtFecha, gdFecSis, True) = True Then
       cmdRendir.Enabled = True
       cmdRendir.SetFocus
    Else
        txtFecha = gdFecSis
    End If
End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdRendir.Enabled = True
    cmdRendir.SetFocus
End If
End Sub

Private Sub VMTpoRend_MenuItemClick(MenuNumber As Long, MenuItem As Long)
cmdRendir.Enabled = False

fraExacta.Visible = False
FraCtaIFPagadora.Visible = False

txtBuscaEntidad.Text = ""
lblIFNombre = ""
lblCtaDesc = ""
lnMenuMumber = MenuNumber
lnMenuItem = MenuItem
Select Case MenuNumber
    Case 1 ' exacta
        Select Case MenuItem
            Case 1
                fraExacta.Visible = True
        End Select
    '***Comentado por ELRO el 20120425, según OYP-RFC005-2012 y OYP-RFC016-2012
    'Case 2 ' Egresos
    '    If lbLoad Then Exit Sub
    '    Select Case MenuItem
    '        Case 1    'Efectivo
    '            cmdRendir_Click
    '        Case 2    'Cheque
    '            FraCtaIFPagadora.Visible = True
    '        Case 3    'Orden Pago
    '            cmdRendir_Click
    '        Case 4    'Carta
    '            FraCtaIFPagadora.Visible = True
    '        Case 5    'Nota de Abono
    '            lsDocTpo = TpoDocNotaCargo
    '            cmdRendir_Click
    '        Case 6
    '            If CCur(Abs(lblSaldo)) >= 1 Then
    '                MsgBox "Saldo no válido para realizar este tipo de Operación. Monto debe ser menor que 1", vbInformation, "Aviso"
    '                Exit Sub
    '            Else
    '                Me.cmdRendir.Enabled = True
    '                Me.txtMovDesc.SetFocus
    '            End If
    '    End Select
    '***Fin Comentado por ELRO*******************************
    Case 3 'Ingresos
        If lbLoad Then Exit Sub
        Select Case MenuItem
            Case 1    'Efectivo
                cmdRendir_Click
            Case 2    'Cheque
                cmdRendir_Click
            Case 3    'Orden Pago
                cmdRendir_Click
            Case 4    'Ingreso por Ventanilla
                cmdRendir_Click
            Case 5    'Nota de Cargo
                cmdRendir_Click
            Case 6
                If CCur(Abs(lblSaldo)) >= 1 Then
                    MsgBox "Saldo no válido para realizar este tipo de Operación. Monto debe ser menor que 1", vbInformation, "Aviso"
                    Exit Sub
                Else
                    Me.cmdRendir.Enabled = True
                    Me.txtMovDesc.SetFocus
                End If
        End Select
End Select
End Sub

Private Sub VMTpoRend_ValidaMenuItem(pnItem As Long, Cancel As Boolean)
Select Case CCur(lnSaldo)
    Case Is = 0
        VMTpoRend.MenusMax = 1
        VMTpoRend.MenuCur = 1
    '***Comentado por ELRO el 20120425, según OYP-RFC005-2012 OYP-RFC016-2012
    'Case Is < 0
    '    If pnItem = 1 Then Cancel = False
    '    VMTpoRend.MenusMax = 2
    '    VMTpoRend.MenuCur = 2
    '***Fin Comentado por ELRO*******************************
    Case Is > 0
        If pnItem = 1 Or pnItem = 2 Then Cancel = False
        VMTpoRend.MenusMax = 3
        VMTpoRend.MenuCur = 3
End Select

End Sub
Function ValidaDatos() As Boolean
ValidaDatos = True
cmdRendir.Enabled = False
Select Case lnMenuMumber
    Case 1
    Case 2
        Select Case lnMenuItem
            Case 1    'Efectivo
            Case 2, 4     ' Carta    'Cheque
                If Len(Trim(txtBuscaEntidad)) = 0 Or lblIFNombre = "" Then
                    MsgBox "Cuenta de Institución Financiera no Ingresada", vbInformation, "Aviso"
                    ValidaDatos = False
                    txtBuscaEntidad.SetFocus
                    Exit Function
                End If
            Case 3    'Orden Pago
            Case 5    'Nota de Cargo
        End Select
    Case 3
        Select Case lnMenuItem
            Case 1    'Efectivo
                
                
            Case 2    'Cheque
                'registro de cheque
            Case 3    'Orden Pago
                
            Case 4    'Nota de Cargo
                
        End Select
End Select
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Descripción o glosa de Operacion no Ingresada ", vbInformation, "Aviso"
    ValidaDatos = False
    txtMovDesc.SetFocus
    Exit Function
End If
cmdRendir.Enabled = True
End Function
Public Property Get vbOk() As Boolean
    vbOk = lbOk
End Property

Public Property Let vbOk(ByVal vNewValue As Boolean)
lbOk = vNewValue
End Property

Private Sub VMTpoRendAge_MenuItemClick(MenuNumber As Long, MenuItem As Long)
cmdRendir.Enabled = False

fraExacta.Visible = False
FraCtaIFPagadora.Visible = False

txtBuscaEntidad.Text = ""
lblIFNombre = ""
lblCtaDesc = ""
lnMenuMumber = MenuNumber
lnMenuItem = MenuItem
Select Case MenuNumber
    Case 1 ' exacta
        Select Case MenuItem
            Case 1
                fraExacta.Visible = True
        End Select
    Case 2 ' Egresos
        If lbLoad Then Exit Sub
        Select Case MenuItem
            Case 1    'Orden Pago
                lnMenuItem = 3
                cmdRendir_Click
            Case 2    'Nota de Abono
                lnMenuItem = 5
               lsDocTpo = TpoDocNotaCargo
                cmdRendir_Click
            Case 3
                lnMenuItem = 6
                If CCur(Abs(lblSaldo)) >= 1 Then
                    MsgBox "Saldo no válido para realizar este tipo de Operación. Monto debe ser menor que 1", vbInformation, "Aviso"
                    Exit Sub
                Else
                    Me.cmdRendir.Enabled = True
                    Me.txtMovDesc.SetFocus
                End If
        End Select
    Case 3 'Ingresos
        If lbLoad Then Exit Sub
        Select Case MenuItem
            Case 1    'Ingreso por Ventanilla
                lnMenuItem = 4
                cmdRendir_Click
            Case 2
                lnMenuItem = 6
                If CCur(Abs(lblSaldo)) >= 1 Then
                    MsgBox "Saldo no válido para realizar este tipo de Operación. Monto debe ser menor que 1", vbInformation, "Aviso"
                    Exit Sub
                Else
                    Me.cmdRendir.Enabled = True
                    Me.txtMovDesc.SetFocus
                End If
        End Select
End Select

End Sub

