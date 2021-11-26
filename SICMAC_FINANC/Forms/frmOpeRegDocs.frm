VERSION 5.00
Begin VB.Form frmOpeRegDocs 
   Caption         =   "A Rendir Cuenta: Regularización: Registro de Documentos"
   ClientHeight    =   6030
   ClientLeft      =   1470
   ClientTop       =   1530
   ClientWidth     =   8055
   ForeColor       =   &H00400000&
   Icon            =   "frmOpeRegDocs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6045
      TabIndex        =   27
      Top             =   5580
      Width           =   1770
   End
   Begin VB.CommandButton cmdAsiento 
      Caption         =   "&Asiento Contable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1905
      TabIndex        =   26
      ToolTipText     =   "Imprimir Documentos Sustentatorios"
      Top             =   5580
      Width           =   1770
   End
   Begin VB.CommandButton cmdImprime 
      Caption         =   "&Planilla de Rendición"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   135
      TabIndex        =   16
      ToolTipText     =   "Imprimir Documentos Sustentatorios"
      Top             =   5580
      Width           =   1770
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
      Left            =   120
      TabIndex        =   8
      Top             =   60
      Width           =   1905
      Begin VB.Label txtRecNro 
         Alignment       =   2  'Center
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
         Left            =   435
         TabIndex        =   18
         Top             =   225
         Width           =   1350
      End
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
         Left            =   120
         TabIndex        =   13
         Top             =   285
         Width           =   255
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
      Left            =   2070
      TabIndex        =   7
      Top             =   60
      Width           =   5745
      Begin VB.Label txtDocFecha 
         Alignment       =   2  'Center
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
         Left            =   4455
         TabIndex        =   21
         Top             =   255
         Width           =   1185
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
         Left            =   2730
         TabIndex        =   20
         Top             =   255
         Width           =   1080
      End
      Begin VB.Label txtDocNro 
         Alignment       =   2  'Center
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
         Left            =   495
         TabIndex        =   19
         Top             =   225
         Width           =   1635
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
         Left            =   2235
         TabIndex        =   14
         Top             =   285
         Width           =   435
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   3945
         TabIndex        =   10
         Top             =   300
         Width           =   555
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
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   255
      End
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
      Height          =   1275
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   7695
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
         Left            =   180
         TabIndex        =   30
         Top             =   525
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
         Left            =   1710
         TabIndex        =   29
         Top             =   510
         Width           =   1050
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
         Left            =   2790
         TabIndex        =   28
         Top             =   510
         Width           =   4830
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
         Height          =   315
         Left            =   2790
         TabIndex        =   25
         Top             =   855
         Width           =   4815
      End
      Begin VB.Label txtAgeDes 
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
         Left            =   2790
         TabIndex        =   24
         Top             =   165
         Width           =   4830
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
         Height          =   315
         Left            =   1215
         TabIndex        =   23
         Top             =   855
         Width           =   1545
      End
      Begin VB.Label txtAgeCod 
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
         Left            =   1710
         TabIndex        =   22
         Top             =   165
         Width           =   1050
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
         Left            =   180
         TabIndex        =   5
         Top             =   885
         Width           =   600
      End
      Begin VB.Label label5 
         Caption         =   "Area Funcional"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   4
         Top             =   210
         Width           =   1125
      End
   End
   Begin VB.Frame fraDoc 
      Caption         =   "&Documentos"
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
      Height          =   3405
      Left            =   135
      TabIndex        =   6
      Top             =   2100
      Width           =   7695
      Begin Sicmact.FlexEdit fgDocSust 
         Height          =   2460
         Left            =   120
         TabIndex        =   17
         Top             =   210
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   4339
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "-Tipo-Número-Proveedor-Fecha-Monto-cMovNro-cTpoDoc-Glosa"
         EncabezadosAnchos=   "300-500-1200-3000-1000-1200-0-0-3000"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-L-L-C-R-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-2-0-0-0"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   300
         RowHeight0      =   285
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6030
         TabIndex        =   2
         Top             =   3000
         Width           =   1210
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
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
         Left            =   1365
         TabIndex        =   1
         Top             =   2790
         Width           =   1200
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "A&gregar"
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
         Left            =   180
         TabIndex        =   0
         Top             =   2790
         Width           =   1200
      End
      Begin VB.TextBox txtTotReg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Left            =   6030
         TabIndex        =   11
         Top             =   2700
         Width           =   1210
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   4890
         X2              =   7260
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5010
         TabIndex        =   15
         Top             =   3030
         Width           =   1035
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total    "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5010
         TabIndex        =   12
         Top             =   2730
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000C&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   4890
         Top             =   2670
         Width           =   2385
      End
   End
End
Attribute VB_Name = "frmOpeRegDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbCaja As Boolean
Dim lnTipoArendir As ArendirTipo
Dim lSalir As Boolean

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
Dim oNArendir As NARendir
Dim lnFaseArendir As ARendirFases
Dim lsAreaCh As String
Dim lsAgeCh As String
Dim lnNroPro As Integer
Dim lbMuestraArendirRendido As Boolean
Dim TmpSaldo As Currency
Dim fsGlosa As String '***Agregado por ELRO el 20130221, según SATI INC1301300007

Public Sub Inicio(ByVal pnFaseArendir As ARendirFases, ByVal pnTipoArendir As ArendirTipo, ByVal pbCaja As Boolean, _
                ByVal psNroArendir As String, ByVal psNroDoc As String, ByVal psFechaDoc As String, _
                ByVal psPersCod As String, ByVal psPersNomb As String, ByVal psAreaCod As String, _
                ByVal psAreaDesc As String, ByVal psAgeCod As String, ByVal psAgeDesc As String, ByVal psDescDoc As String, ByVal psMovNroAtencion As String, _
                ByVal pnImporte As Currency, ByVal psCtaContARendir As String, ByVal psCtaContPendiente As String, _
                ByVal pnSaldo As Currency, ByVal psMovNroSol As String, Optional psAreaCh As String = "", _
                Optional psAgeCh As String = "", Optional pnNroProc As Integer = 0, Optional pbMuestraArendirRendido As Boolean = False, _
                Optional psGlosa As String = "")
                '***Parametro psGlosa agregado por ELRO el 20130221, según SATI INC1301300007
                
                
lsAreaCh = psAreaCh
lsAgeCh = psAgeCh
lnNroPro = pnNroProc
               
lnFaseArendir = pnFaseArendir
lsNroArendir = psNroArendir
lsNroDoc = psNroDoc
lsFechaDoc = psFechaDoc
lsPersCod = psPersCod
lsPersNomb = psPersNomb
lsAreaCod = psAreaCod
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
lbCaja = pbCaja
lnTipoArendir = pnTipoArendir
lbMuestraArendirRendido = pbMuestraArendirRendido
fsGlosa = psGlosa '***Agregado por ELRO el 20130221, según SATI INC1301300007

Me.Show 1
End Sub

Private Sub cmdAgregar_Click()
Dim nPos As Integer
Dim lnFila As Integer

TmpSaldo = txtSaldo.Text

If Val(txtSaldo) = 0 Then
    If MsgBox("Arendir ha sido sustentado y no posee Saldo. Desea Continuar??", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
End If
frmOpeDocChica.Inicio lnFaseArendir, lnTipoArendir, Trim(txtAgeCod), Trim(txtAgeDes), _
                        lblAgeCod, lblAgeDesc, Trim(txtPerCod), Trim(txtPerNom), _
                        txtDocFecha, lsMovNroAtencion, lsMovNroSol, txtRecNro, TmpSaldo, lbCaja, _
                        CCur(txtRecImporte), lsAreaCh, lsAgeCh, lnNroPro, fsGlosa
                        '***Parametro fsGlosa agregado por ELRO el 20130221, según SATI INC1301300007
If frmOpeDocChica.lOk Then
    Me.fgDocSust.AdicionaFila
    lnFila = Me.fgDocSust.Row
    fgDocSust.TextMatrix(lnFila, 1) = frmOpeDocChica.DocAbrev
    fgDocSust.TextMatrix(lnFila, 2) = frmOpeDocChica.DocNro '  gsDocNro
    fgDocSust.TextMatrix(lnFila, 3) = frmOpeDocChica.Proveedor ' gcEntiOrig
    fgDocSust.TextMatrix(lnFila, 4) = frmOpeDocChica.FechaDoc '  gdFecha
    fgDocSust.TextMatrix(lnFila, 5) = Format(frmOpeDocChica.ImporteDoc * IIf(frmOpeDocChica.DocAbrev = "NC", -1, 1), gsFormatoNumeroView)   '  Format(gnImporte, gsFormatoNumeroView)
    fgDocSust.TextMatrix(lnFila, 6) = frmOpeDocChica.MovNroSust
    fgDocSust.TextMatrix(lnFila, 8) = frmOpeDocChica.MovDesc
    txtTotReg = Format(Val(Format(txtTotReg, gsFormatoNumeroDato)) + (frmOpeDocChica.ImporteDoc * IIf(frmOpeDocChica.DocAbrev = "NC", -1, 1)), gsFormatoNumeroView)
    txtSaldo = Format(Val(Format(txtRecImporte, gsFormatoNumeroDato)) - Val(Format(txtTotReg, gsFormatoNumeroDato)), gsFormatoNumeroView)
    If Val(txtSaldo) < 0 Then
        txtSaldo.ForeColor = &HFF&
    Else
        txtSaldo.ForeColor = &HC00000
    End If
End If
End Sub

Private Sub cmdAsiento_Click()
Dim lsTexto As String
Dim oContImp As NContImprimir
Set oContImp = New NContImprimir
If txtRecImporte = txtSaldo Then
   MsgBox "No se ha declarado ningún Gasto...!", vbInformation, "Aviso"
   Exit Sub
End If
Me.MousePointer = 11
'***Modificado por ELRO el 20120425, según OYP-RFC005-2012
'lsTexto = oContImp.ImprimeAsientoSustArendir(lsMovNroAtencion, gnColPage, gnLinPage, gsNomCmac, gdFecSis, lsCtaContARendir, gsSimbolo, lnTipoArendir)
lsTexto = oContImp.ImprimeAsientoSustArendir(lsMovNroAtencion, gnColPage, gnLinPage, gsNomCmac, gdFecSis, lsCtaContArendir, gsSimbolo, lnTipoArendir, gsOpeCod)
'***Fin Modificado por ELRO*******************************
lsTexto = lsTexto + oImpresora.gPrnSaltoPagina + oContImp.ImprimeDetDocCtaCont(lsMovNroAtencion, gnColPage, gnLinPage, lsCtaContArendir, lsCtaContPendiente, gsSimbolo, lnTipoArendir)
lsTexto = lsTexto & oImpresora.gPrnSaltoPagina & lsTexto
Me.MousePointer = 0
EnviaPrevio lsTexto, Me.Caption, gnLinPage
End Sub

Private Sub cmdCerrar_Click()
'gnImporte = Val(Format(txtSaldo, gsFormatoNumerodato))
lnSaldo = CCur(txtSaldo)
Unload Me
End Sub
Private Sub cmdEliminar_Click()
Dim lsMovNroAct As String
If fgDocSust.TextMatrix(fgDocSust.Row, 0) = "" Then
   Exit Sub
End If
If MsgBox(" ¿ Seguro de Eliminar Documento Sustentatorio ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
    lsMovNroAct = GeneraMovNroActualiza(Format(Now, gsFormatoFechaView), gsCodUser, gsCodCMAC, gsCodAge)
    
    If oNArendir.EliminaMovDocSust(fgDocSust.TextMatrix(fgDocSust.Row, 6), lsMovNroSol, lnTipoArendir, _
                         CCur(fgDocSust.TextMatrix(fgDocSust.Row, 5)), CCur(txtSaldo.Text), lsMovNroAct) = 0 Then
        'MADM 20110606
        oNArendir.EliminaGastoRecupera fgDocSust.TextMatrix(fgDocSust.Row, 6)
        'END MADM
        fgDocSust.EliminaFila fgDocSust.Row
        txtTotReg = Format(fgDocSust.SumaRow(5), gsFormatoNumeroView)
        txtSaldo = Format(lnImporte - CCur(txtTotReg), "#,#0.00")
        If Val(txtSaldo) < 0 Then
            txtSaldo.ForeColor = &HFF&
        Else
            txtSaldo.ForeColor = &HC00000
        End If
    End If
End If
End Sub

Private Sub cmdImprime_Click()
Dim lsTexto As String
Dim oContImp As NContImprimir
Set oContImp = New NContImprimir
If txtRecImporte = txtSaldo Then
   MsgBox "No se ha declarado ningún Gasto...!", vbInformation, "Aviso"
   Exit Sub
End If
Me.MousePointer = 11
'***Modificado por ELRO el 20120425, según OYP-RFC005-2012
'lsTexto = oContImp.ImprimePlanillaRendicion(lsMovNroAtencion, gnColPage, gnLinPage, gsNomCmac, gdFecSis, lsCtaContARendir, lsCtaContPendiente, IIf(Mid(gsOpeCod, 3, 1) = gMonedaNacional, gcMN, gcME), Mid(gsOpeCod, 3, 1), lnTipoArendir)
lsTexto = oContImp.ImprimePlanillaRendicion(lsMovNroAtencion, gnColPage, gnLinPage, gsNomCmac, gdFecSis, lsCtaContArendir, lsCtaContPendiente, IIf(Mid(gsOpeCod, 3, 1) = gMonedaNacional, gcMN, gcME), Mid(gsOpeCod, 3, 1), lnTipoArendir, gsOpeCod)
'***Modificado por ELRO***********************************
Me.MousePointer = 0
EnviaPrevio lsTexto, Me.Caption, gnLinPage
Set oContImp = Nothing
End Sub



Private Sub Form_Load()
CentraForm Me
fraDocEmit.Caption = lsDescDoc
txtPerCod = lsPersCod
txtPerNom = lsPersNomb
txtDocFecha = lsFechaDoc
txtRecImporte = Format(lnImporte, gsFormatoNumeroView)
txtAgeCod = lsAreaCod
txtAgeDes = lsAreaDesc
lblAgeDesc = lsAgeDesc
lblAgeCod = lsAgeCod
 
If lnTipoArendir = gArendirTipoViaticos Then
   txtRecNro = lsNroDoc
Else
   txtDocNro = lsNroDoc
   txtRecNro = lsNroArendir
End If

If gsSimbolo = gcMN Then
   Label6.Caption = "Total    " & gcMN
   Label8.Caption = "Saldo    " & gcMN
Else
   Label6.Caption = "Total    " & gcME
   Label8.Caption = "Saldo    " & gcME
End If
If lnTipoArendir = gArendirTipoCajaChica Then
   Frame3.Visible = False
   fraDocEmit.Caption = "Recibo de A rendir Caja Chica"
   fraDocEmit.Left = Frame1.Left
   cmdImprime.Visible = False
   cmdAsiento.Visible = False
End If
If lnFaseArendir = ArendirRendicion Or lbMuestraArendirRendido Then
   cmdAgregar.Enabled = False
   cmdEliminar.Enabled = False
End If

Set oNArendir = New NARendir
CargaDocSustentados
End Sub
Private Sub CargaDocSustentados()
Dim rs As ADODB.Recordset
Dim lnFila As Long
Set rs = New Recordset
Set rs = oNArendir.GetDocSustentariosArendir(lsMovNroAtencion, lsCtaContArendir, lsCtaContPendiente)
Do While Not rs.EOF
    fgDocSust.AdicionaFila
    lnFila = fgDocSust.Row
       
    fgDocSust.TextMatrix(lnFila, 1) = rs!cDocAbrev
    fgDocSust.TextMatrix(lnFila, 2) = rs!cDocNro
    fgDocSust.TextMatrix(lnFila, 3) = PstaNombre(rs!cPersNombre)
    fgDocSust.TextMatrix(lnFila, 4) = Format(rs!dDocFecha, "dd/mm/yyyy")
    fgDocSust.TextMatrix(lnFila, 5) = Format(rs!nDocImporte, gsFormatoNumeroView)
    fgDocSust.TextMatrix(lnFila, 6) = rs!cMovNro
    fgDocSust.TextMatrix(lnFila, 7) = rs!nDocTpo
    fgDocSust.TextMatrix(lnFila, 8) = rs!cMovDesc
    rs.MoveNext
Loop
rs.Close: Set rs = Nothing
txtTotReg = Format(fgDocSust.SumaRow(5), "#,#0.00")
txtSaldo = Format(lnImporte - CCur(txtTotReg), "#,#0.00")
If Val(txtSaldo) < 0 Then
    txtSaldo.ForeColor = &HFF&
Else
    txtSaldo.ForeColor = &HC00000
End If
End Sub

Private Sub txtSaldo_GotFocus()
cmdCerrar.SetFocus
End Sub


