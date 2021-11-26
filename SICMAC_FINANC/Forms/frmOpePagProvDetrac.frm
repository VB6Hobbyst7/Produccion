VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5F774E03-DB36-4DFC-AAC4-D35DC9379F2F}#1.1#0"; "VertMenu.ocx"
Begin VB.Form frmOpePagProvDetrac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  RETIROS PARA PAGOS DE DETRACCIONES"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   Icon            =   "frmOpePagProvDetrac.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   Begin VertMenu.VerticalMenu vFormaPago 
      Height          =   6135
      Left            =   0
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   10821
      MenuCaption1    =   "Forma Pago"
      MenuItemsMax1   =   3
      MenuItemIcon11  =   "frmOpePagProvDetrac.frx":08CA
      MenuItemCaption11=   "Efectivo"
      MenuItemKey11   =   "1"
      MenuItemIcon12  =   "frmOpePagProvDetrac.frx":0BE4
      MenuItemCaption12=   "Carta"
      MenuItemKey12   =   "2"
      MenuItemIcon13  =   "frmOpePagProvDetrac.frx":0EFE
      MenuItemCaption13=   "Cheque"
      MenuItemKey13   =   "3"
   End
   Begin VB.TextBox txtNroConstancia 
      Height          =   315
      Left            =   1440
      TabIndex        =   20
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CheckBox chkImpresion 
      Caption         =   "Imprimir Doc. en Soles"
      Height          =   540
      Left            =   9975
      TabIndex        =   19
      Top             =   105
      Value           =   1  'Checked
      Width           =   1110
   End
   Begin VB.CommandButton cmdSinDetra 
      Caption         =   "&Obviar Detracción"
      Height          =   345
      Left            =   7170
      TabIndex        =   15
      Top             =   5100
      Width           =   1545
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "&Rechazar"
      Height          =   345
      Left            =   8850
      TabIndex        =   16
      Top             =   5550
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdDoc 
      Caption         =   "&Emitir"
      Height          =   345
      Left            =   8775
      TabIndex        =   14
      Top             =   5100
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   9930
      TabIndex        =   17
      Top             =   5100
      Width           =   1095
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   690
      Left            =   1290
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4110
      Width           =   9765
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
      Left            =   1290
      TabIndex        =   3
      Top             =   4890
      Width           =   5820
      Begin Sicmact.TxtBuscar txtBuscaEntidad 
         Height          =   345
         Left            =   120
         TabIndex        =   4
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
         TabIndex        =   5
         Top             =   210
         Width           =   3705
      End
   End
   Begin VB.Frame FraFechaMov 
      Height          =   495
      Left            =   6060
      TabIndex        =   0
      Top             =   150
      Width           =   2040
      Begin MSMask.MaskEdBox txtFechaMov 
         Height          =   300
         Left            =   870
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   180
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imgRec 
      Left            =   4710
      Top             =   5010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpePagProvDetrac.frx":1218
            Key             =   "recibo"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtxtAsiento 
      Height          =   315
      Left            =   10140
      TabIndex        =   13
      Top             =   5370
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmOpePagProvDetrac.frx":1312
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
   Begin Sicmact.FlexEdit fg 
      Height          =   3315
      Left            =   1290
      TabIndex        =   18
      Top             =   750
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   5847
      Cols0           =   16
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmOpePagProvDetrac.frx":1392
      EncabezadosAnchos=   "400-0-500-2100-1140-4000-0-1250-0-0-0-0-0-0-1400-1250"
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
      ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-C-L-L-R-L-C-C-C-C-C-C-R"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "Nro."
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
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
      Height          =   660
      Left            =   1290
      TabIndex        =   7
      Top             =   30
      Width           =   8505
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
         Left            =   6930
         TabIndex        =   8
         Top             =   180
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   315
         Left            =   750
         TabIndex        =   9
         Top             =   240
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
         Left            =   2640
         TabIndex        =   10
         Top             =   240
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   105
         TabIndex        =   12
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         Height          =   195
         Left            =   2040
         TabIndex        =   11
         Top             =   285
         Width           =   510
      End
   End
   Begin MSMask.MaskEdBox txtFechaConst 
      Height          =   315
      Left            =   3600
      TabIndex        =   24
      Top             =   5880
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Constancia "
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
      Left            =   3600
      TabIndex        =   23
      Top             =   5640
      Width           =   1875
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Nro. de Constancia "
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
      Left            =   1200
      TabIndex        =   22
      Top             =   0
      Width           =   1710
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nro. de Constancia "
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
      Left            =   1440
      TabIndex        =   21
      Top             =   5640
      Width           =   1710
   End
End
Attribute VB_Name = "frmOpePagProvDetrac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCtaContDebeB As String
Dim lsCtaContDebeS As String
Dim lsDocs      As String
Dim lsFileCarta As String
Dim lsDocTpo    As String

Dim rs As New ADODB.Recordset
Dim lSalir As Boolean
Dim lMN As Boolean
Dim lTransActiva As Boolean
Dim ContSalOp As Integer
Dim sFileOP   As String
Dim sFileVE   As String
Dim sFileNA   As String
Dim sFileVNA  As String
Dim fs     As Scripting.FileSystemObject
Dim oBarra As clsProgressBar
Dim cCtaEmbargo As String
Dim cCtaFielCump As String

Dim lsPersId As String

Dim lsCtaITFD As String
Dim lsCtaITFH As String
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Function ValidaInterfaz() As Boolean
Dim oCon As New DConecta
Dim reg As New ADODB.Recordset

 

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

If Mid(gsOpeCod, 3, 1) = "1" And chkImpresion.value = 0 Then
    oCon.AbreConexion
    Set reg = oCon.CargaRecordSet("Select count(*) nCantidad From Mov M Inner Join MovCta MC On M.nMovNro=MC.nMovNro Where MC.cCtaContCod like '__2%' And M.cMovNro='" & fg.TextMatrix(fg.row, 9) & "'")
    If reg.BOF Then
    Else
        If reg!nCantidad > 0 Then
            Set reg = Nothing
            oCon.CierraConexion
            MsgBox "Debe efectuar la detracción desde su correspondiente opción en dólares", vbInformation, "Aviso"
            Exit Function
        End If
    End If
    Set reg = Nothing
    oCon.CierraConexion
End If

ValidaInterfaz = True
End Function
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

Dim oDocPago      As clsDocPago
Dim lsSubCuentaIF As String
Dim lsPersCod     As String
Dim lsDocNRo      As String
Dim lsMovNro      As String
Dim lsDocumento   As String
Dim lsOpeCod      As String
Dim lsCtaBanco    As String
Dim lsCtaContHaber As String
Dim lsPersCodIf    As String
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

'
Dim lsCtaContDebeBMNE As String
Dim lnMontoMNE1 As Double
Dim lnMontoMNE2 As Double

lsCtaContDebeBMNE = ""
'If Mid(gsOpeCod, 3, 1) = "2" Then
'    lsCtaContDebeBMNE = Left(lsCtaContDebeB, 2) & "" & Mid(lsCtaContDebeB, 4, Len(lsCtaContDebeB) - 3)
'End If
lsCtaContDebeBMNE = lsCtaContDebeB


Dim oNCaja As nCajaGeneral
Set oNCaja = New nCajaGeneral

'Retencion
Dim oConst As NConstSistemas
Set oConst = New NConstSistemas
Dim oImpuesto As DImpuesto
Set oImpuesto = New DImpuesto

Dim lbBitReten As Boolean
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
   
'Detracciones
Dim oCont As New DConecta
Dim RegT As New ADODB.Recordset
Dim bDolares As Boolean
Dim nDolares1 As Double

bDolares = False
nDolares1 = 0

If Mid(gsOpeCod, 3, 1) = "1" Then
    oCont.AbreConexion
    Set RegT = oCont.CargaRecordSet("Select count(*) nCantidad From Mov M Inner Join MovCta MC On M.nMovNro=MC.nMovNro Where MC.cCtaContCod like '__2%' And M.cMovNro='" & fg.TextMatrix(fg.row, 9) & "'")
    If RegT.BOF Then
    Else
        If RegT!nCantidad > 0 Then
            Set RegT = Nothing
            oCont.CierraConexion
            bDolares = True
            
        End If
    End If
    Set RegT = Nothing
    oCont.CierraConexion
End If
'Fin Detracciones
   
lbBitReten = IIf(oConst.LeeConstSistema(gConstSistBitRetencion6Porcent) = 1, True, False)

On Error GoTo NoGrabo
If txtFechaMov.Enabled Then
    If Not ValidaFechaContab(txtFechaMov, gdFecSis) Then
        Exit Sub
    End If
End If
If ValidaInterfaz = False Then Exit Sub

'Add by gitu 05-03-2009
'If Not gsOpeCod = gOpeCGOpeBancosRetCtasBancosDetracMN Or Not gsOpeCod = gOpeCGOpeBancosRetCtasBancosDetracME Then
'    lsDocTpo = TpoDocCheque
'End If
'end gitu

If lsDocTpo = "" Then
    MsgBox "Seleccione Forma de Pago", vbInformation, "Aviso"
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
lsPersCodIf = Mid(txtBuscaEntidad, 4, 13)
lsEntidadOrig = oDCtaIF.NombreIF(lsPersCodIf)
lsSubCuentaIF = oCtasIF.SubCuentaIF(lsPersCodIf)

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
lnMontoMNE1 = 0: lnMontoMNE2 = 0

lsCabeImpre = " DOCUMENTOS PAGADOS : "
nDocs = 0

For K = 1 To fg.Rows - 1
   If fg.TextMatrix(K, 2) = "." Then
         
'        k = fg.row
      
      nDocs = nDocs + 1
      
      '*** PEAC 20110303
      If nDocs > 1 Then
        MsgBox "Los Pagos se realizan de forma individual por documento, por favor realice el pago de un documento y despues continue con otro documento.", vbOKOnly + vbInformation, "Aviso"
        fg.SetFocus
        Exit Sub
      End If
      '*** FIN PEAC
      
      If Not lsPersCod = "" Then
         If Not lsPersCod = fg.TextMatrix(fg.row, 8) Then
            MsgBox "No se puede hacer Pago a Proveedores Diferentes", vbInformation, "¡Aviso!"
            fg.SetFocus
            Exit Sub
         End If
      End If
      lsPersCod = fg.TextMatrix(fg.row, 8)
      lsPersNombre = fg.TextMatrix(fg.row, 5)
      lsCabeImpre = lsCabeImpre & oImpresora.gPrnCondensadaON & fg.TextMatrix(K, 3) & Space(5) & oImpresora.gPrnCondensadaOFF
      If nDocs Mod 4 = 0 Then
         lsCabeImpre = lsCabeImpre & oImpresora.gPrnSaltoLinea & Space(22)
      End If
      
      
      
      If lsCtaContDebeB = fg.TextMatrix(K, 13) Then
         lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 7))
         
         'Agregado Detraccion
         If Mid(gsOpeCod, 3, 1) = "1" And bDolares = True And chkImpresion.value = 1 Then
            nDolares1 = nDolares1 + CCur(fg.TextMatrix(K, 15))
         End If
         'Fin Detraccion
         
      End If
   
        '
        If Mid(gsOpeCod, 3, 1) = "2" Then
            If lsCtaContDebeBMNE = fg.TextMatrix(K, 13) Then
                lnMontoMNE1 = lnMontoMNE1 + CCur(fg.TextMatrix(K, 7))
                If Mid(gsOpeCod, 3, 1) = "1" And bDolares = True And chkImpresion.value = 1 Then
                    lnMontoMNE2 = lnMontoMNE2 + CCur(fg.TextMatrix(K, 15))
                Else
                    If gsOpeCod = gOpeCGOpeBancosRetFielCumplimientoME Then
                       lnMontoMNE2 = lnMontoMNE2 + CCur(fg.TextMatrix(K, 15))
                    End If
                End If
            End If
            If gsOpeCod = gOpeCGOpeBancosRetFielCumplimientoME Then
                lsCtaContDebeBMNE = lsCtaContDebeB
            Else
               lsCtaContDebeBMNE = Mid(lsCtaContDebeB, 1, 2) & "1" & Mid(lsCtaContDebeB, 4, Len(lsCtaContDebeB) - 3)
            End If
        End If
        
   End If
Next
If lnImporteB + lnImporteS = 0 Then
   
    '
    If lnMontoMNE1 > 0 Then
        
        lnImporteB = lnMontoMNE1
        
    Else
    '
        MsgBox "No se Seleccionó Comprobantes para Pagar!", vbInformation, "¡Aviso!"
        fg.SetFocus
        Exit Sub
    End If
Else
    If nDolares1 > 0 And chkImpresion.value = 1 Then
        lnImporteB = nDolares1
    End If
End If

If lbBitReten Then
    lbBCAR = VerifBCAR(lsPersCod)
    lsCtaReten = oConst.LeeConstSistema(gConstSistCtaRetencion6Porcent)
    lnTasaImp = oImpuesto.CargaImpuesto(lsCtaReten)!nImpTasa
    lnIngresos = oNCaja.GetMontoIngresoRetencion(lsPersCod, Left(Format(gdFecSis, gsFormatoMovFecha), 6), True)
    lnRetencion = oNCaja.GetMontoIngresoRetencion(lsPersCod, Left(Format(gdFecSis, gsFormatoMovFecha), 6), False)
    lnTopeRetencion = oConst.LeeConstSistema(gConstSistTopeRetencion6Porcent)
    
    If Not lbBCAR Then
        If lMN Then
            lnRetAct = (lnImporteB + lnImporteS) + lnIngresos
        Else
            lnRetAct = Round((lnImporteB + lnImporteS) * gnTipCambioPonderado, 2) + lnIngresos
        End If
        If lnRetAct <= lnTopeRetencion Then
           lnRetAct = 0
        Else
           lnRetAct = Round(lnRetAct * (lnTasaImp / 100), 2) - lnRetencion
           
           If lMN Then
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

    If lMN Then
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

If lsDocTpo = "-1" Then
    frmArendirEfectivo.inicio 0, fg.TextMatrix(fg.row, 12), Mid(gsOpeCod, 3, 1), "", lnImporteB + lnImporteS - IIf(lMN, lnRetAct, lnRetActME), lsPersCod, lsPersNombre, ArendirRendicion, "Nro.Doc.:"
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
    lsDocTpo = TpoDocNotaAbono
    
    frmNotaCargoAbono.inicio lsDocTpo, lnImporteB + lnImporteS - IIf(lMN, lnRetAct, lnRetActME), gdFecSis, txtMovDesc, gsOpeCod, False, lsPersCod, lsPersNombre, , , lnITFValor
    If frmNotaCargoAbono.vbOk Then
        lsDocNRo = frmNotaCargoAbono.NroNotaCA
        txtMovDesc = frmNotaCargoAbono.Glosa
        lsDocumento = frmNotaCargoAbono.NotaCargoAbono
        lsPersNombre = frmNotaCargoAbono.PersNombre
        lsPersDireccion = frmNotaCargoAbono.PersDireccion
        lsUbigeo = frmNotaCargoAbono.PersUbigeo
        lsCuentaAho = frmNotaCargoAbono.CuentaAhoNro
        lsFecha = frmNotaCargoAbono.FechaNotaCA
'        lsDocumento = oImp.ImprimeNotaCargoAbono(lsDocNRo, txtMovDesc, CCur(frmNotaCargoAbono.Monto), _
'                            lsPersNombre, lsPersDireccion, lsUbigeo, gdFecSis, Mid(gsOpeCod, 3, 1), lsCuentaAho, lsDocTpo, gsNomAge, gsCodUser)
         lsDocumento = oImp.ImprimeNotaAbono(lsFecha, lnImporteB + lnImporteS - IIf(lMN, lnRetAct, lnRetActME), txtMovDesc, lsCuentaAho, lsPersNombre)
        lbGrabaOpeNegocio = MsgBox(" ¿ Desea que se realice Abono en Cuenta del Proveedor ? ", vbQuestion + vbYesNo, "¡Confirmacion!") = vbYes
        If lbGrabaOpeNegocio Then
            Dim oDis As New NRHProcesosCierre
            lsCadBol = oDis.ImprimeBoletaCad(CDate(lsFecha), "ABONO CAJA GENERAL", "Depósito CAJA GENERAL*Nro." & lsDocNRo, "", lnImporteB + lnImporteS - IIf(lMN, lnRetAct, lnRetActME), lsPersNombre, lsCuentaAho, "", 0, 0, "Nota Abono", 0, 0, False, False, , , , True, , , , False, gsNomAge) & oImpresora.gPrnSaltoPagina
        End If
    Else
        Exit Sub
    End If
Else
    If lsDocTpo = TpoDocCheque Then
       lsDocVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1), gsCodAge)
        If Mid(gsOpeCod, 3, 1) = "2" Then
            If chkImpresion.value = 0 Then
                'oDocPago.InicioCheque lsDocNRo, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeDesc, gsGlosa, lnMontoMNE1, gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, True, gsCodAge
                oDocPago.InicioCheque lsDocNRo, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeDesc, gsGlosa, lnMontoMNE1, gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, True, gsCodAge, , , lsTpoIf, lsPersCodIf, lsCtaBanco 'EJVG20121130
            Else
                'oDocPago.InicioCheque lsDocNRo, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeDesc, gsGlosa, lnMontoMNE2, gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, True, gsCodAge
                oDocPago.InicioCheque lsDocNRo, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeDesc, gsGlosa, lnMontoMNE2, gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, True, gsCodAge, , , lsTpoIf, lsPersCodIf, lsCtaBanco 'EJVG20121130
            End If
        Else
            'oDocPago.InicioCheque lsDocNRo, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeDesc, gsGlosa, lnImporteB + lnImporteS - IIf(lMN, lnRetAct, lnRetActME), gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, True, gsCodAge
            oDocPago.InicioCheque lsDocNRo, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeDesc, gsGlosa, lnImporteB + lnImporteS - IIf(lMN, lnRetAct, lnRetActME), gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, True, gsCodAge, , , lsTpoIf, lsPersCodIf, lsCtaBanco 'EJVG20121130
        End If
    End If
    If lsDocTpo = TpoDocOrdenPago Then
       lsDocVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1), gsCodAge)
       oDocPago.InicioOrdenPago lsDocNRo, True, lsPersCod, gsOpeCod, lsPersNombre, gsOpeCod, gsGlosa, lnImporteB + lnImporteS - IIf(lMN, lnRetAct, lnRetActME), gdFecSis, lsDocVoucher, True, gsCodAge
    End If
    If lsDocTpo = TpoDocCarta Then
       oDocPago.InicioCarta lsDocNRo, lsPersCod, gsOpeCod, gsOpeCod, gsGlosa, lsFileCarta, lnImporteB + lnImporteS - lnRetAct, gdFecSis, lsEntidadOrig, lsCtaEntidadOrig, lsPersNombre, "", lsMovNro, gnMgDer, gnMgIzq, gnMgSup
    End If
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
       Exit Sub
    End If
End If
lsOpeCod = oOpe.EmiteOpeDoc(Mid(gsOpeCod, 1, 6), lsDocTpo)
If lsOpeCod = "" Then
   MsgBox "No se asignó Documentos de Referencia a Operación de Pago", vbInformation, "Aviso"
   Exit Sub
End If

If lsDocTpo = TpoDocOrdenPago Then
   lsCtaContHaber = oOpe.EmiteOpeCta(lsOpeCod, "H", , gsCodArea, ObjCMACAgenciaArea)
Else
   lsCtaContHaber = oOpe.EmiteOpeCta(lsOpeCod, "H", , txtBuscaEntidad, ObjEntidadesFinancieras)
End If

''*** PEAC 20111114
'If Left(lsCtaContHaber, 1) = "X" Then
'   MsgBox "En la Operación xxx la cuenta contable xxx que pertenece a xxx, No está definida, por favor comunique a contabilidad.", vbInformation, "Aviso"
'   Exit Sub
'End If
''*** FIN PEAC

If lsCtaContDebeB = "" Or lsCtaContHaber = "" Then
   MsgBox "Cuentas Contables no determinadas Correctamente." & oImpresora.gPrnSaltoLinea & "consulte con Sistemas", vbInformation, "Aviso"
   Exit Sub
End If

If Mid(gsOpeCod, 3, 1) = "2" Then
    lsCtaContDebeB = lsCtaContDebeBMNE
End If

If (gsOpeCod = gOpeCGOpeBancosRetCtasBancosDetracMN Or gsOpeCod = gOpeCGOpeBancosRetCtasBancosDetracME) And (txtFechaConst.Text = "" Or txtNroConstancia.Text = "") Then
   MsgBox "Debe Ingresar Fecha y Nro de Constancia de Detracción" & oImpresora.gPrnSaltoLinea & "consulte con Sistemas", vbInformation, "Aviso"
   Exit Sub
End If
 
If MsgBox("Desea Grabar la Información", vbYesNo + vbQuestion, "Aviso") = vbYes Then
   cmdDoc.Enabled = False
   lsMovNro = oNContFunc.GeneraMovNro(txtFechaMov, Right(gsCodAge, 2), gsCodUser)
   lsCtaDiferencia = oOpe.EmiteOpeCta(lsOpeCod, IIf(lnMontoDif > 0, "H", "D"), "2")
   
   
   If gsOpeCod = gOpeCGOpeBancosRetCtasBancosDetracMN Or gsOpeCod = gOpeCGOpeBancosRetCtasBancosDetracME Then
        If oNCaja.GrabaPagoProveedorDetraccion(lsMovNro, lsOpeCod, txtMovDesc, lsCtaContDebeB, lsCtaContDebeS, _
                                   lsCtaContHaber, lnImporteB, lnImporteS, lsPersCod, lsTpoIf, lsPersCodIf, lsCtaBanco, _
                                   rsBilletaje, lsDocTpo, lsDocNRo, Format(CDate(lsFecha), gsFormatoFecha), lsDocVoucher, fg.GetRsNew, lsCuentaAho, lsCtaDiferencia, lnMontoDif, gbBitCentral, lbGrabaOpeNegocio, lnRetAct, lsCtaITFD, lsCtaITFH, gnImpITF, True, , , , , , , , , txtFechaConst.Text, txtNroConstancia.Text) = 0 Then
                                     
                              
            ImprimeAsientoContable lsMovNro, lsDocVoucher, lsDocTpo, lsDocumento, lbEfectivo, False, txtMovDesc, lsPersCod, lnImporteB + lnImporteS - IIf(lMN, lnRetAct, lnRetActME), , , , 1, , "17", , lsCabeImpre, lsCadBol
            K = 1
            Do While K < fg.Rows
               If fg.TextMatrix(K, 2) = "." Then
                  fg.EliminaFila K
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
                                     
    Else
        If gsOpeCod = gOpeCGOpeBancosRetFielCumplimientoME Then
           lnImporteB = lnMontoMNE2
        End If
    
        If oNCaja.GrabaPagoProveedor(lsMovNro, lsOpeCod, txtMovDesc, lsCtaContDebeB, lsCtaContDebeS, _
                                   lsCtaContHaber, lnImporteB, lnImporteS, lsPersCod, lsTpoIf, lsPersCodIf, lsCtaBanco, _
                                   rsBilletaje, lsDocTpo, lsDocNRo, Format(CDate(lsFecha), gsFormatoFecha), lsDocVoucher, fg.GetRsNew, lsCuentaAho, lsCtaDiferencia, lnMontoDif, gbBitCentral, lbGrabaOpeNegocio, lnRetAct, lsCtaITFD, lsCtaITFH, gnImpITF, True) = 0 Then
                                   
                
            ImprimeAsientoContable lsMovNro, lsDocVoucher, lsDocTpo, lsDocumento, lbEfectivo, False, txtMovDesc, lsPersCod, lnImporteB + lnImporteS - IIf(lMN, lnRetAct, lnRetActME), , , , 1, , "17", , lsCabeImpre, lsCadBol
            K = 1
            Do While K < fg.Rows
               If fg.TextMatrix(K, 2) = "." Then
                  fg.EliminaFila K
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
    End If
   cmdDoc.Enabled = True

    'ARLO20170217
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Grabo Operación"
    Set objPista = Nothing
    '****

End If
Exit Sub

NoGrabo:
  MsgBox TextErr(Err.Description), vbInformation, "Error de Actualización"
  cmdDoc.Enabled = True
End Sub


Private Sub cmdProcesar_Click()
On Error GoTo ErrProcesar
   If CDate(txtFechaDel) > CDate(txtFechaAl) Then
      MsgBox "Fecha de Inicio no puede ser Mayor que Fecha final", vbInformation, "Aviso"
      Exit Sub
   End If
   cmdProcesar.Enabled = False
   CargaProveedores
   cmdProcesar.Enabled = True
Exit Sub
ErrProcesar:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub cmdRechazar_Click()
On Error GoTo ErrSave
If Len(fg.TextMatrix(fg.row, 0)) > 0 Then
    If MsgBox(" ¿ Seguro de Rechazar comprobante de Proveedor ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
       Dim oMov As New DMov
       oMov.EliminaMov fg.TextMatrix(fg.row, 9)
       Set oMov = Nothing
       fg.EliminaFila fg.row
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
 

Private Sub cmdSinDetra_Click()
Dim lnMovNro As Long
Dim oMov As New DMov

If Len(Trim(fg.TextMatrix(fg.row, 9))) > 0 Then
    Detrae
    cmdProcesar_Click
End If
End Sub

Private Function Detrae()
Dim lnMovNro As Long

Dim oMov As New DMov
 
If MsgBox("Desea obviar la detracción?" & Chr(10) & Chr(10) & "Esto sólo se deberá usar cuando no se haya provisionado el comprobante" & Chr(10) & "y ya se haya efectuado el depósito en el banco", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    lnMovNro = oMov.GetnMovNro(fg.TextMatrix(fg.row, 9))
     
    oMov.BeginTrans
    oMov.InsertaMovDetra lnMovNro, 2, gsCodUser
     
    oMov.ActualizaMovPendientesRend lnMovNro, "29180799", 0, False
    oMov.ActualizaMovPendientesRend lnMovNro, "29280799", 0, False
    
    oMov.CommitTrans
    Set oMov = Nothing
    
    'ARLO20170217
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Obvio Detracción "
    Set objPista = Nothing
    '****
    
    cmdProcesar_Click
End If
End Function

Private Sub fg_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    Dim lsPersIdentif As String
    Dim lnI As Long
    
    If fg.TextMatrix(pnRow, 2) = "." Then
        lsPersIdentif = fg.TextMatrix(pnRow, 8)
        
        For lnI = 1 To Me.fg.Rows - 1
            If fg.TextMatrix(lnI, 2) = "." And fg.TextMatrix(lnI, 8) <> lsPersIdentif Then
               fg.TextMatrix(lnI, 2) = ""
            End If
        Next lnI
        
        fg.row = pnRow
    End If
    
End Sub

Private Sub fg_RowColChange()
txtMovDesc = fg.TextMatrix(fg.row, 6)
End Sub

Private Sub Form_Activate()
If lSalir Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim nTipCambio As Currency
On Error GoTo LoadErr
CentraForm Me

lSalir = False

txtFechaDel = DateAdd("D", -30, gdFecSis)
txtFechaAl = gdFecSis
txtFechaMov = gdFecSis
lsDocTpo = "-1"
ContSalOp = 0

gnDocTpo = 0
gsDocNro = ""
gsGlosa = ""


Dim oOpe As New DOperacion
lsCtaContDebeB = oOpe.EmiteOpeCta(gsOpeCod, "D", "0")
lsCtaContDebeS = oOpe.EmiteOpeCta(gsOpeCod, "D", "1")

Set rs = oOpe.CargaOpeDoc(gsOpeCod, , OpeDocMetDigitado)
lsDocs = RSMuestraLista(rs, 1)
Set oOpe = Nothing
RSClose rs

'If gsOpeCod = OpeCGOpeProvRechazo Then
'   fraFecha.Left = fraFecha.Left - 1020
'   fraFecha.Width = fraFecha.Width + 1020
'   fg.Left = fg.Left - 1020
'   fg.Width = fg.Width + 1020
'   txtMovDesc.Left = txtMovDesc.Left - 1020
'   txtMovDesc.Width = txtMovDesc.Width + 1020
'   cmdDoc.Visible = False
'   CmdImprimir.Visible = False
'   cmdRechazar.Visible = True
'   Me.Height = 5750
'   Exit Sub
'End If

If gsOpeCod = gOpeCGOpeBancosRetCtasBancosEmbargoMN Or gsOpeCod = gOpeCGOpeBancosRetCtasBancosEmbargoME Then
   frmOpePagProvDetrac.Caption = "OPERACION PAGO DE EMBARGO SUNAT" & Space(2) & IIf(Mid(gsOpeCod, 3, 1) = 1, "MN", "ME")
   chkImpresion.value = 0
   chkImpresion.Enabled = False
End If

If gsOpeCod = gOpeCGOpeBancosRetFielCumplimientoMN Or gsOpeCod = gOpeCGOpeBancosRetFielCumplimientoME Then
   frmOpePagProvDetrac.Caption = "OPERACION PAGO DE FIEL CUMPLIMIENTO" & Space(2) & IIf(Mid(gsOpeCod, 3, 1) = 1, "MN", "ME")
   chkImpresion.value = 0
   chkImpresion.Enabled = False
End If

If gsOpeCod = gOpeCGOpeBancosRetCtasBancosDetracMN Or gsOpeCod = gOpeCGOpeBancosRetCtasBancosDetracME Then
   chkImpresion.value = 0
   chkImpresion.Enabled = False
   vFormaPago.Visible = True 'Add by Gitu 05-03-2009
End If

lMN = IIf(Mid(gsOpeCod, 3, 1) = Moneda.gMonedaExtranjera, False, True)

'lsFileCarta = App.path & "\" & gsDirPlantillas & gsOpeCod & ".TXT"
txtBuscaEntidad.psRaiz = "Cuentas de Instituciones Financieras"
Set oOpe = New DOperacion
Dim nOpeCod As String

If gsOpeCod = 402581 Or gsOpeCod = 402582 Then
    nOpeCod = Mid(gsOpeCod, 1, 2) & "1" & Mid(gsOpeCod, 4, Len(Trim(gsOpeCod)))
    txtBuscaEntidad.rs = oOpe.GetRsOpeObj(nOpeCod, "1")  '  oDCtaIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), gTpoIFBanco + gTpoCtaIFCtaCte + gTpoCtaIFCtaAho)
Else
    txtBuscaEntidad.rs = oOpe.GetRsOpeObj(gsOpeCod, "1")  '  oDCtaIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), gTpoIFBanco + gTpoCtaIFCtaCte + gTpoCtaIFCtaAho)
End If
Set oOpe = Nothing

lsDocTpo = TpoDocCheque
txtFechaMov.Enabled = True

lsCtaITFD = oOpe.EmiteOpeCta(gsOpeCod, "D", 2)
lsCtaITFH = oOpe.EmiteOpeCta(gsOpeCod, "H", 2)

If gsOpeCod = 401581 Or gsOpeCod = 402581 Then
   Me.Height = 6800
Else
   Me.Height = 6120
End If

Exit Sub
LoadErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"

End Sub
Private Sub CargaProveedores()
Dim rs As New ADODB.Recordset
Dim nItem As Long
Dim cCtaDetraTemp As String
Dim cCtaDetraTempMNE As String
Dim cCtaEmbargoTempMNE As String
Dim cCtaFielCumpTempMNE As String
Dim lsCtaContEmbargoMNE As String
Dim lsCtaFielCumpMNE As String
Dim lsCtaContDebeBMNE As String
Dim oDCaja As New DCajaGeneral

On Error GoTo ErrCargaProveedores
Set oBarra = New clsProgressBar

fg.Clear
fg.Rows = 2
fg.FormaCabecera

Dim oCon As NConstSistemas
Set oCon = New NConstSistemas

Select Case gsOpeCod
         
    Case gOpeCGOpeBancosRetCtasBancosEmbargoMN, gOpeCGOpeBancosRetCtasBancosEmbargoME

        'Embargo SUNAT
            
        cCtaEmbargo = oCon.LeeConstSistema(168)
        cCtaEmbargoTempMNE = cCtaEmbargo
        lsCtaContEmbargoMNE = cCtaEmbargo
    
        Set rs = oDCaja.GetDatosProvisionesProveedoresNuevo("'" & lsCtaContEmbargoMNE & "'", lsDocs, txtFechaDel, txtFechaAl, , 2, cCtaEmbargoTempMNE, , , , , gsOpeCod)
         
    Case gOpeCGOpeBancosRetFielCumplimientoMN, gOpeCGOpeBancosRetFielCumplimientoME
        ' Fiel Cumplimiento
    
        cCtaFielCump = oCon.LeeConstSistema(172)
        If Mid(gsOpeCod, 3, 1) = 1 Then
           cCtaFielCump = Mid(cCtaFielCump, 1, 2) & "1" & Mid(cCtaFielCump, 4, Len(Trim(cCtaFielCump)))
        Else
           cCtaFielCump = Mid(cCtaFielCump, 1, 2) & "2" & Mid(cCtaFielCump, 4, Len(Trim(cCtaFielCump)))
        End If
        
        cCtaFielCumpTempMNE = cCtaFielCump
        lsCtaFielCumpMNE = cCtaFielCump
           
        Set rs = oDCaja.GetDatosProvisionesProveedoresNuevo("'" & lsCtaFielCumpMNE & "'", lsDocs, txtFechaDel, txtFechaAl, , 2, cCtaFielCumpTempMNE, , , , , gsOpeCod)
         
    Case gOpeCGOpeBancosRetCtasBancosDetracMN, gOpeCGOpeBancosRetCtasBancosDetracME
         'Detraccion
         cCtaDetraTemp = cCtaDetraccionProvision
         cCtaDetraTempMNE = cCtaDetraTemp
         lsCtaContDebeBMNE = cCtaDetraTemp
           
         Set rs = oDCaja.GetDatosProvisionesProveedoresNuevo("'" & lsCtaContDebeBMNE & "'", lsDocs, txtFechaDel, txtFechaAl, , 2, cCtaDetraTempMNE, , , , , gsOpeCod)

End Select

Set oDCaja = Nothing
Set oCon = Nothing
If rs.EOF Then
   RSClose rs
   cmdProcesar.Enabled = True
   MsgBox "No existen Comprobantes Pendientes", vbInformation, "Aviso"
   Exit Sub
End If

Do While Not rs.EOF
   fg.AdicionaFila
   nItem = fg.row
   fg.TextMatrix(nItem, 1) = nItem
   fg.TextMatrix(nItem, 3) = Mid(rs!cDocAbrev & Space(3), 1, 3) & " " & rs!cDocNro
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
   fg.TextMatrix(nItem, 15) = Format(rs!nMovImporteME, gsFormatoNumeroView)
   rs.MoveNext
Loop
RSClose rs
fg.row = 1
txtMovDesc = fg.TextMatrix(1, 6)
Exit Sub
ErrCargaProveedores:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub txtBuscaEntidad_EmiteDatos()
Dim oCtaIf As NCajaCtaIF
Set oCtaIf = New NCajaCtaIF
lblCtaDesc = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaEntidad, 18, 10)) + " " + txtBuscaEntidad.psDescripcion
If txtBuscaEntidad <> "" Then
   If gsOpeCod = 401581 Or gsOpeCod = 402581 Then
      txtNroConstancia.SetFocus
   Else
      cmdDoc.SetFocus
   End If
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

Private Sub txtFechaConst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdDoc.SetFocus
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

Private Sub txtNroConstancia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtFechaConst.SetFocus
End If
End Sub
'Add By gitu 05-03-2009
Private Sub vFormaPago_MenuItemClick(MenuNumber As Long, MenuItem As Long)
    Select Case MenuItem
        Case 1: lsDocTpo = "-1"
        Case 2: lsDocTpo = TpoDocCarta
        Case 3: lsDocTpo = TpoDocCheque
    End Select
    cmdDoc.SetFocus
End Sub
