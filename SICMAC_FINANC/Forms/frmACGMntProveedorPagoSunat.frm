VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5F774E03-DB36-4DFC-AAC4-D35DC9379F2F}#1.1#0"; "VertMenu.ocx"
Begin VB.Form frmACGMntProveedorPagoSunat 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   Icon            =   "frmACGMntProveedorPagoSunat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7185
      TabIndex        =   16
      Top             =   4890
      Value           =   1  'Checked
      Width           =   1605
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   420
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   4200
      Width           =   9795
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
      Left            =   1170
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   5820
      Begin Sicmact.TxtBuscar txtBuscaEntidad 
         Height          =   345
         Left            =   120
         TabIndex        =   13
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
         EnabledText     =   0   'False
      End
      Begin VB.Label lblCtaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2010
         TabIndex        =   14
         Top             =   210
         Width           =   3705
      End
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
      Left            =   1200
      TabIndex        =   3
      Top             =   0
      Width           =   9825
      Begin VB.Frame FraFechaMov 
         Height          =   495
         Left            =   5790
         TabIndex        =   9
         Top             =   120
         Width           =   2040
         Begin MSMask.MaskEdBox txtFechaMov 
            Height          =   300
            Left            =   870
            TabIndex        =   10
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
            TabIndex        =   11
            Top             =   180
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Buscar"
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
         Left            =   8190
         TabIndex        =   8
         Top             =   210
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   315
         Left            =   750
         TabIndex        =   4
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
         TabIndex        =   5
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         Height          =   195
         Left            =   2040
         TabIndex        =   7
         Top             =   285
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9975
      TabIndex        =   1
      Top             =   4845
      Width           =   1095
   End
   Begin VB.CommandButton cmdEmitir 
      Caption         =   "&Emitir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8850
      TabIndex        =   0
      Top             =   4845
      Width           =   1095
   End
   Begin Sicmact.FlexEdit fg 
      Height          =   3375
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   5953
      Cols0           =   21
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmACGMntProveedorPagoSunat.frx":08CA
      EncabezadosAnchos=   "500-0-500-2100-1140-4000-0-1250-0-0-0-0-0-0-0-1500-2000-3000-0-2000-1200"
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
      ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-L-L-L-R-L-C-C-C-L-L-L-R-R-R-C-L-L"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-2-2-2-0-0-4"
      TextArray0      =   "Nro."
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   495
      RowHeight0      =   360
      ForeColorFixed  =   -2147483630
   End
   Begin VertMenu.VerticalMenu vFormaPago 
      Height          =   5230
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   1000
      _ExtentX        =   1773
      _ExtentY        =   9234
      MenuCaption1    =   "Forma Pago"
      MenuItemsMax1   =   3
      MenuItemIcon11  =   "frmACGMntProveedorPagoSunat.frx":0993
      MenuItemCaption11=   "Efectivo"
      MenuItemIcon12  =   "frmACGMntProveedorPagoSunat.frx":0CAD
      MenuItemCaption12=   "Carta"
      MenuItemIcon13  =   "frmACGMntProveedorPagoSunat.frx":0FC7
      MenuItemCaption13=   "Cheque"
   End
End
Attribute VB_Name = "frmACGMntProveedorPagoSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCtaContDebeB As String
Dim lsCtaContDebeS As String
Dim lsCtaContDebeBME As String
Dim lsCtaContDebeSME As String

Dim lsCtaContDebeRH As String
Dim lsCtaContDebeRHJ As String
Dim lsCtaContDebeSegu As String
Dim lsCtaContDebePagoVarios As String

Dim lsCtaContDebeRHME As String
Dim lsCtaContDebeRHJME As String
Dim lsCtaContDebeSeguME As String
Dim lsCtaContDebePagoVariosME As String

'Created 30-07-2008 Gitu
Dim lsDocTpo As String
Dim lMN As Boolean
Dim lsFileCarta As String
'Fin Gitu
Dim objPista As COMManejador.Pista

Dim lsCtaITFD As String
Dim lsCtaITFH As String
Private Sub cmdEmitir_Click()
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
Dim lnImporteRH    As Currency
Dim lnImporteRHJ    As Currency
Dim lnImporteSEGU    As Currency
Dim lnImportePAGOVARIOS    As Currency

Dim lnImporteBME    As Currency
Dim lnImporteSME    As Currency
Dim lnImporteRHME    As Currency
Dim lnImporteRHJME    As Currency
Dim lnImporteSEGUME    As Currency
Dim lnImportePAGOVARIOSME    As Currency

Dim lnImporteDolaresAcum As Currency
Dim lnImporteDolaresAcumS As Currency
Dim lnImporteDolaresAcumRH As Currency
Dim lnImporteDolaresAcumRHJ As Currency
Dim lnImporteDolaresAcumSegu As Currency
Dim lnImporteDolaresAcumPagoVarios As Currency

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

Dim oNCaja As DCajaGeneral 'nCajaGeneral
Set oNCaja = New DCajaGeneral

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
Dim nItem As Integer
Dim lsEquipo As String
Set oPrevio = New clsPrevioFinan ' clsPrevio

On Error GoTo ErrNoGrabo

If lsDocTpo = "" Then
    MsgBox "Seleccione Forma de Pago", vbInformation, "Aviso"
    Exit Sub
End If

Set oNContFunc = New NContFunciones
Set oDCtaIF = New DCajaCtasIF
Set oCtasIF = New NCajaCtaIF
Set oDocPago = New clsDocPago


    'lsDocVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1), gsCodAge)
    
    lsCtaEntidadOrig = Trim(lblCtaDesc)
    lsTpoIf = Mid(txtBuscaEntidad, 1, 2)
    lsCtaBanco = Mid(txtBuscaEntidad, 18, Len(Me.txtBuscaEntidad))
    lsPersCodIf = Mid(txtBuscaEntidad, 4, 13)
    lsEntidadOrig = oDCtaIF.NombreIF(lsPersCodIf)
    lsSubCuentaIF = oCtasIF.SubCuentaIF(lsPersCodIf)
    
    lsDocVoucher = ""
    lsDocNRo = ""
    lbEfectivo = False
    lsCadBol = ""
    lsCtaContHaber = ""
    lsPersCod = ""
    lnImporteB = 0: lnImporteS = 0
    lnImporteBME = 0: lnImporteSME = 0
    lnImporteDolaresAcum = 0
    lnImporteDolaresAcumS = 0
    lsCabeImpre = " DOCUMENTOS PAGADOS : "
    nDocs = 0
    For K = 1 To fg.Rows - 1
       If fg.TextMatrix(K, 2) = "." Then
          nItem = K
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
          lsCabeImpre = lsCabeImpre & oImpresora.gPrnCondensadaON & fg.TextMatrix(K, 3) & Space(5) & oImpresora.gPrnCondensadaOFF
          If nDocs Mod 4 = 0 Then
             lsCabeImpre = lsCabeImpre & oImpresora.gPrnSaltoLinea & Space(22)
          End If
          If lsCtaContDebeB = fg.TextMatrix(K, 13) Then
             lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 16))
          End If
          If lsCtaContDebeBME = fg.TextMatrix(K, 13) Then
             lnImporteBME = lnImporteBME + CCur(fg.TextMatrix(K, 16))
             lnImporteDolaresAcum = lnImporteDolaresAcum + CCur(fg.TextMatrix(K, 18))
          End If
          If lsCtaContDebeS = fg.TextMatrix(K, 13) Then
             lnImporteS = lnImporteS + CCur(fg.TextMatrix(K, 16))
          End If
          If lsCtaContDebeSME = fg.TextMatrix(K, 13) Then
             lnImporteSME = lnImporteSME + CCur(fg.TextMatrix(K, 16))
             lnImporteDolaresAcumS = lnImporteDolaresAcumS + CCur(fg.TextMatrix(K, 18))
          End If
          

        If lsCtaContDebeRH = fg.TextMatrix(K, 13) Then
           lnImporteRH = lnImporteRH + CCur(fg.TextMatrix(K, 16))
        End If
        If lsCtaContDebeRHME = fg.TextMatrix(K, 13) Then
           lnImporteRHME = lnImporteRHME + CCur(fg.TextMatrix(K, 16))
           lnImporteDolaresAcumRH = lnImporteDolaresAcumRH + CCur(fg.TextMatrix(K, 18))
        End If
                
        If lsCtaContDebeRHJ = fg.TextMatrix(K, 13) Then
           lnImporteRHJ = lnImporteRHJ + CCur(fg.TextMatrix(K, 16))
        End If
        If lsCtaContDebeRHJME = fg.TextMatrix(K, 13) Then
           lnImporteRHJME = lnImporteRHJME + CCur(fg.TextMatrix(K, 16))
           lnImporteDolaresAcumRHJ = lnImporteDolaresAcumRHJ + CCur(fg.TextMatrix(K, 18))
        End If
                       
        If lsCtaContDebeSegu = fg.TextMatrix(K, 13) Then
           lnImporteSEGU = lnImporteSEGU + CCur(fg.TextMatrix(K, 16))
        End If
        If lsCtaContDebeSeguME = fg.TextMatrix(K, 13) Then
           lnImporteSEGUME = lnImporteSEGUME + CCur(fg.TextMatrix(K, 16))
           lnImporteDolaresAcumSegu = lnImporteDolaresAcumSegu + CCur(fg.TextMatrix(K, 18))
        End If
                               
        If lsCtaContDebePagoVarios = fg.TextMatrix(K, 13) Then
           lnImportePAGOVARIOS = lnImportePAGOVARIOS + CCur(fg.TextMatrix(K, 16))
        End If
        If lsCtaContDebePagoVariosME = fg.TextMatrix(K, 13) Then
           lnImportePAGOVARIOSME = lnImportePAGOVARIOSME + CCur(fg.TextMatrix(K, 16))
           lnImporteDolaresAcumPagoVarios = lnImporteDolaresAcumPagoVarios + CCur(fg.TextMatrix(K, 18))
        End If
       End If
    Next
    If lnImporteB + lnImporteS + lnImporteBME + lnImporteSME = 0 Then
       MsgBox "No se Seleccionó Comprobantes para Pagar!", vbInformation, "¡Aviso!"
       fg.SetFocus
       Exit Sub
    End If
    'Add by GITU 04-08-08
    If lsDocTpo = "-1" Then
        frmArendirEfectivo.Inicio 0, fg.TextMatrix(fg.Row, 12), Mid(gsOpeCod, 3, 1), "", lnImporteB + lnImporteS + lnImporteSEGU - IIf(lMN, lnRetAct, lnRetActME), lsPersCod, lsPersNombre, ArendirRendicion, "Nro.Doc.:"
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
    Else
        If lsDocTpo = TpoDocCheque Then
            lsDocVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1))
            'oDocPago.InicioCheque lsDocNRo, True, Mid(Me.txtBuscaEntidad, 4, 13), gsOpeCod, "SUNAT/BANCO DE LA NACION", gsOpeDesc, Me.txtMovDesc, lnImporteB + lnImporteS + lnImporteBME + lnImporteSME + lnImporteRH + lnImporteRHJ + lnImporteSEGU + lnImportePAGOVARIOS + lnImporteRHME + lnImporteRHJME + lnImporteSEGUME + lnImportePAGOVARIOSME, gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, True, gsCodAge
            oDocPago.InicioCheque lsDocNRo, True, Mid(Me.txtBuscaEntidad, 4, 13), gsOpeCod, "SUNAT/BANCO DE LA NACION", gsOpeDesc, Me.txtMovDesc, lnImporteB + lnImporteS + lnImporteBME + lnImporteSME + lnImporteRH + lnImporteRHJ + lnImporteSEGU + lnImportePAGOVARIOS + lnImporteRHME + lnImporteRHJME + lnImporteSEGUME + lnImportePAGOVARIOSME, gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, True, gsCodAge, , , lsTpoIf, lsPersCodIf, lsCtaBanco 'EJVG20121130
        End If
        If lsDocTpo = TpoDocCarta Then
            oDocPago.InicioCarta lsDocNRo, lsPersCod, gsOpeCod, gsOpeCod, gsGlosa, lsFileCarta, lnImporteB + lnImporteS + lnImporteBME + lnImporteSME + lnImporteRH + lnImporteRHJ + lnImporteSEGU - lnImportePAGOVARIOS + lnImporteRHME + lnImporteRHJME + lnImporteSEGUME + lnImportePAGOVARIOSME + lnRetAct, gdFecSis, lsEntidadOrig, lsCtaEntidadOrig, lsPersNombre, "", lsMovNro, gnMgDer, gnMgIzq, gnMgSup
        End If
        If oDocPago.vbOk Then    'Se ingresó dato de Cheque u Orden de Pago
            lsOpeCod = gsOpeCod
            lsFecha = oDocPago.vdFechaDoc
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
    'End Gitu
    
    lsOpeCod = gsOpeCod '"421120"
    If lsOpeCod = "" Then
       MsgBox "No se asignó Documentos de Referencia a Operación de Pago", vbInformation, "Aviso"
       Exit Sub
    End If

    Set oOpe = New DOperacion
    lsCtaContHaber = oOpe.EmiteOpeCta(lsOpeCod, "H", , txtBuscaEntidad, ObjEntidadesFinancieras)

    
    If lsCtaContDebeB = "" Or lsCtaContDebeS = "" Or lsCtaContHaber = "" Then
       MsgBox "Cuentas Contables no determinadas Correctamente." & oImpresora.gPrnSaltoLinea & "consulte con Sistemas", vbInformation, "Aviso"
       Exit Sub
    End If
 lsEquipo = GetMaquinaUsuario

If MsgBox("Desea Grabar la Información", vbYesNo + vbQuestion, "Aviso") = vbYes Then
   cmdEmitir.Enabled = False
   lsMovNro = oNContFunc.GeneraMovNro(txtFechaMov, Right(gsCodAge, 2), gsCodUser)
   If oNCaja.GrabaPagoProvSUNAT(lsMovNro, gsOpeCod, "PAGO SUNAT " & Me.txtMovDesc, lsCtaContDebeB, lsCtaContDebeS, lsCtaContDebeRH, lsCtaContDebeRH, lsCtaContDebeSegu, lsCtaContDebePagoVarios, lsCtaContDebeBME, lsCtaContDebeSME, lsCtaContDebeRHME, lsCtaContDebeRHJME, lsCtaContDebeSeguME, lsCtaContDebePagoVariosME, _
                                   lsCtaContHaber, lnImporteB, lnImporteS, lnImporteRH, lnImporteRHJ, lnImporteSEGU, lnImportePAGOVARIOS, lnImporteBME, lnImporteSME, lnImporteRHME, lnImporteRHJME, lnImporteSEGUME, lnImportePAGOVARIOSME, lsPersCod, lsTpoIf, lsPersCodIf, lsCtaBanco, _
                                   Me.fg.Recordset, TpoDocCheque, lsDocNRo, Format(CDate(lsFecha), gsFormatoFecha), lsDocVoucher, fg.GetRsNew, "", lsCtaDiferencia, lnMontoDif, gbBitCentral, , , lnITFValor, , , lnImporteDolaresAcum, lnImporteDolaresAcumS, lnImporteDolaresAcumRH, lnImporteDolaresAcumRHJ, lnImporteDolaresAcumSegu, lnImporteDolaresAcumPagoVarios, _
                                   lsCtaITFD, lsCtaITFH, gnImpITF, IIf(chkAfectoITF.value = 1, True, False)) = 0 Then
      
      ImprimeAsientoContable lsMovNro, lsDocVoucher, TpoDocCheque, lsDocumento, lbEfectivo, False, txtMovDesc, lsPersCod, lnImporteB + lnImporteS + lnImporteRH + lnImporteRHJ + lnImporteSEGU + lnImportePAGOVARIOS, , , , 1, , "17", , ""
      K = 1
      objPista.InsertarPista lsOpeCod, lsMovNro, gsCodPersUser, lsEquipo, "1", "Pago SUNAT"
      Do While K < fg.Rows
         If fg.TextMatrix(K, 2) = "." Then
            fg.EliminaFila K
         Else
            K = K + 1
         End If
      Loop
      cmdEmitir.Enabled = True
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
      lsDocNRo = ""
      lsDocVoucher = ""
      lsDocumento = ""
      txtBuscaEntidad = ""
      lblCtaDesc = ""
   End If
   cmdEmitir.Enabled = True
End If

Exit Sub

ErrNoGrabo:
  MsgBox TextErr(Err.Description), vbInformation, "Error de Actualización"
  cmdEmitir.Enabled = True
End Sub

Private Sub cmdProcesar_Click()
Dim rs As New ADODB.Recordset
Dim nItem As Long
Dim nTieneDetra As New DMov
Dim bTieneDetra As Boolean
Dim cCtaDetraTemp As String
Dim nCantTempo As Integer
Dim oDCaja As New DCajaGeneral

On Error GoTo ErrCargaProveedores

fg.Clear
fg.Rows = 2
fg.FormaCabecera

'Dim oDCaja As New DCajaGeneral


'cCtaDetraTemp = Mid(cCtaDetraccionProvision, 1, 2) & Mid(gsOpeCod, 3, 1) & Mid(cCtaDetraccionProvision, 4, Len(cCtaDetraccionProvision) - 2)


Set rs = oDCaja.GetDatosPagoProveedoresSUNAT(txtFechaDel, DateAdd("d", 1, CDate(txtFechaAl)), gsOpeCod)
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
    fg.TextMatrix(nItem, 15) = Format(rs!nMovImporteSoles, gsFormatoNumeroView)
    fg.TextMatrix(nItem, 16) = Format(rs!nimportecoactivo, gsFormatoNumeroView)
   'fg.TextMatrix(nItem, 17) = IIf(IsNull(rs!MovPago), "", rs!MovPago)
    If rs!nMovImporteSoles <> rs!nMovImporte Then
        fg.TextMatrix(nItem, 19) = "DOLARES"
        fg.TextMatrix(nItem, 18) = Round(rs!nimportecoactivo / rs!nTpoCambio, 2)
    Else
        fg.TextMatrix(nItem, 19) = "SOLES"
        fg.TextMatrix(nItem, 18) = 0
    End If
    
    rs.MoveNext
Loop
RSClose rs
fg.Row = 1
txtMovDesc = fg.TextMatrix(1, 6)

If nCantTempo > 0 Then
    MsgBox "Existe(n) " & nCantTempo & " registro(s) que no se cargó porque aún falta registrar la detracción", vbInformation, "Aviso"
End If

Exit Sub
ErrCargaProveedores:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

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
        
        fg.Row = pnRow
    End If
End Sub


Private Sub fg_OnRowChange(pnRow As Long, pnCol As Long)
    txtMovDesc = fg.TextMatrix(fg.Row, 6)
End Sub

Private Sub Form_Load()
    Dim oOpe As New DOperacion
    Set oOpe = New DOperacion
    Me.txtFechaDel = DateAdd("m", -1, CDate(gdFecSis))
    Me.txtFechaAl = gdFecSis
    Me.txtFechaMov = gdFecSis
    txtBuscaEntidad.rs = oOpe.GetRsOpeObj(gsOpeCod, "1")
    lsCtaContDebeB = oOpe.EmiteOpeCta(gsOpeCod, "D", "0")
    lsCtaContDebeS = oOpe.EmiteOpeCta(gsOpeCod, "D", "1")
    lsCtaContDebeRH = oOpe.EmiteOpeCta(gsOpeCod, "D", "3") 'RRHH Descuento Planillas
    lsCtaContDebeRHJ = oOpe.EmiteOpeCta(gsOpeCod, "D", "4") 'RRHH Judicial
    lsCtaContDebeSegu = oOpe.EmiteOpeCta(gsOpeCod, "D", "5") 'Seguros
    lsCtaContDebePagoVarios = oOpe.EmiteOpeCta(gsOpeCod, "D", "6") 'Pagos Varios
    
    Set objPista = New COMManejador.Pista
    
    lsCtaContDebeBME = "252601"
    lsCtaContDebeSME = "25260202"
    lsCtaContDebeB = "251601"
    lsCtaContDebeS = "25160202"
    Me.txtFechaMov.Enabled = True
    
    'Add By Gitu 30-07-2008
    lMN = IIf(Mid(gsOpeCod, 3, 1) = Moneda.gMonedaExtranjera, False, True)
    lsFileCarta = App.path & "\" & gsDirPlantillas & gsOpeCod & ".TXT"
    
    'Add by Gitu 16-03-2009
    If gsOpeCod = "421120" Then
        Me.Caption = "PAGO SUNAT"
    End If
    'Fin Gitu
    
    lsCtaITFD = oOpe.EmiteOpeCta(gsOpeCod, "D", 2)
    lsCtaITFH = oOpe.EmiteOpeCta(gsOpeCod, "H", 2)
       
End Sub

Private Sub txtBuscaEntidad_EmiteDatos()
Dim oCtaIf As NCajaCtaIF
Set oCtaIf = New NCajaCtaIF
lblCtaDesc = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaEntidad, 18, 10)) + " " + txtBuscaEntidad.psDescripcion
If txtBuscaEntidad <> "" Then
   Me.cmdEmitir.SetFocus
End If
Set oCtaIf = Nothing
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
    End If
End Sub


Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFechaAl.SetFocus
    End If
End Sub

Private Sub vFormaPago_MenuItemClick(MenuNumber As Long, MenuItem As Long)
    txtFechaMov.Enabled = False
    txtFechaMov = gdFecSis
    Select Case MenuItem
        Case 1: 'Efectivo
                fraEntidad.Visible = False
                fraEntidad.Visible = False
                lsDocTpo = "-1"
                cmdEmitir_Click
        Case 2: ' TpoDocCarta
                fraEntidad.Visible = True
                lsDocTpo = TpoDocCarta
                txtFechaMov.Enabled = True
        Case 3:  'Cheque
                fraEntidad.Visible = True
                lsDocTpo = TpoDocCheque
                txtFechaMov.Enabled = True
    End Select
If fraEntidad.Visible Then
   txtBuscaEntidad.SetFocus
Else
   If cmdEmitir.Visible Then
    cmdEmitir.SetFocus
   End If
End If
End Sub

