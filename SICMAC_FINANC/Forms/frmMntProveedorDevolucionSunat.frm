VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMntProveedorDevolucionSunat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caption"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10440
   Icon            =   "frmMntProveedorDevolucionSunat.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAbonoSUNAT 
      Caption         =   "Abono SUNAT"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Top             =   6960
      Width           =   10215
      Begin VB.ComboBox cboMedioPagoSUNAT 
         Height          =   315
         ItemData        =   "frmMntProveedorDevolucionSunat.frx":030A
         Left            =   1920
         List            =   "frmMntProveedorDevolucionSunat.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
      Begin Sicmact.TxtBuscar txtBuscaEntidad2 
         Height          =   345
         Left            =   1920
         TabIndex        =   26
         Top             =   600
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
      Begin VB.Label lblCtaDesc2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3840
         TabIndex        =   27
         Top             =   600
         Width           =   4425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Medio de Pago"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lblEntidadPagadora2 
         AutoSize        =   -1  'True
         Caption         =   "Entidad Pagadora"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1275
      End
   End
   Begin VB.Frame fraPagosRealizados 
      Caption         =   "Pagos Realizados"
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
      TabIndex        =   0
      Top             =   0
      Width           =   5865
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
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   315
         Left            =   750
         TabIndex        =   2
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   285
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   105
         TabIndex        =   1
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdRegDevolucion 
      Caption         =   "Realizar &Devolución"
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
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Frame fraGlosa 
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
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   8415
      Begin VB.TextBox txtMovDesc 
         Height          =   540
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   8115
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
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
      Left            =   8760
      TabIndex        =   29
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   8760
      TabIndex        =   28
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Registrar"
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
      Left            =   8760
      TabIndex        =   20
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame fraDevolucion 
      Caption         =   "Devolución"
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
      TabIndex        =   10
      Top             =   5040
      Width           =   8385
      Begin VB.TextBox txtNResolucion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   190
         Width           =   3135
      End
      Begin MSMask.MaskEdBox txtFechaDev 
         Height          =   315
         Left            =   6960
         TabIndex        =   14
         Top             =   195
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
      Begin VB.Label lblFechaDevolucion 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Resolución:"
         Height          =   195
         Left            =   5280
         TabIndex        =   13
         Top             =   255
         Width           =   1560
      End
      Begin VB.Label lblNResolucion 
         AutoSize        =   -1  'True
         Caption         =   "Nº Resolución SUNAT:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1665
      End
   End
   Begin Sicmact.FlexEdit fg 
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4895
      Cols0           =   19
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmMntProveedorDevolucionSunat.frx":038E
      EncabezadosAnchos=   "300-1000-1000-1000-1000-0-3000-4500-2500-0-0-0-0-0-0-0-0-0-0"
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
      ColumnasAEditar =   "X-X-X-X-4-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-R-L-L-L-C-C-C-C-C-C-C-C-C-L-R"
      FormatosEdit    =   "0-0-0-0-2-0-0-0-0-0-0-0-0-0-0-0-0-0-2"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame fraAbono 
      Caption         =   "Abono a Proveedor"
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
      Height          =   1140
      Left            =   120
      TabIndex        =   15
      Top             =   5760
      Width           =   10185
      Begin Sicmact.TxtBuscar txtBuscaEntidad 
         Height          =   345
         Left            =   1920
         TabIndex        =   19
         Top             =   600
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
      Begin VB.Label lblMedioPago 
         AutoSize        =   -1  'True
         Caption         =   "Medio de Pago"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lblMedioPag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblEntidadPagadora 
         AutoSize        =   -1  'True
         Caption         =   "Entidad Pagadora"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label lblCtaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3810
         TabIndex        =   21
         Top             =   600
         Width           =   4425
      End
   End
End
Attribute VB_Name = "frmMntProveedorDevolucionSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************
'Nombre : frmMntProveedorDevolucionSunat
'Descripcion:Formulario para Devolucion de pagos realizados a la SUNAT
'Creacion: PASIERS1242014
'*****************************
Option Explicit
Dim lsDocTpo As String
Dim lsFileCarta As String

Dim objPista As COMManejador.Pista
Dim lsCtaITFD As String
Dim lsCtaITFH As String
Dim fsPersCodCMACMaynas As String
Dim fsPersCodBCP As String
Dim fnTpoPago As Integer
Dim lMN As Boolean
Dim lsCtaCajaGeneral As String
Dim gsOpeCodAsiento2 As String
Private Sub cmdCancelar_Click()
    EstablecerDatosInicio
    DesHabilitaControles 0
End Sub
Private Sub EstablecerDatosInicio()
    fnTpoPago = 0
    txtMovDesc.Text = ""
    txtNResolucion.Text = ""
    txtFechaDev = gdFecSis
    lblMedioPag.Caption = ""
    txtBuscaEntidad.Text = ""
    cboMedioPagoSUNAT.ListIndex = -1
    txtBuscaEntidad2.Text = ""
    lblCtaDesc.Caption = ""
    lblCtaDesc2.Caption = ""
    'fg.SetFocus
    cmdRegDevolucion.Enabled = True
    cmdProcesar.Enabled = True
    fg.Enabled = True
    'fg.SetFocus
End Sub
Private Function ValidaInterfaz() As Boolean
    If fnTpoPago = 0 Then
        MsgBox "El sistema no ha podido determinar la forma de devolución.", vbInformation, "Aviso"
            If fg.TextMatrix(fg.row, 1) = "" Then
                cmdProcesar.SetFocus
            Else
                fg.SetFocus
            End If
            ValidaInterfaz = False
            Exit Function
    End If
    If fnTpoPago = gPagoCuentaCMAC Then
        If Len(fg.TextMatrix(fg.row, 12)) = 0 Then
            MsgBox "El proveedor " & fg.TextMatrix(fg.row, 6) & "no tiene configurado una cuenta en Soles para realizar la devolución. Coordine con el Dpto. de Logística.", vbInformation, "Aviso"
            ValidaInterfaz = False
            Exit Function
        End If
    End If
    
    If fg.TextMatrix(fg.row, 4) > fg.TextMatrix(fg.row, 18) Then
        MsgBox "El importe a Devolver al proveedor no puede ser superior al Monto devuelto por la SUNAT.", vbInformation, "Aviso"
        ValidaInterfaz = False
        Exit Function
    End If
    
    If Len(txtMovDesc.Text) = 0 Then
        MsgBox "No se ha ingresado la descripcion de la devolución.", vbInformation, "Aviso"
        txtMovDesc.SetFocus
        ValidaInterfaz = False
        Exit Function
    End If
    If Len(txtNResolucion.Text) = 0 Then
        MsgBox "No se ha ingresado el número de resolución.", vbInformation, "Aviso"
        txtNResolucion.SetFocus
        ValidaInterfaz = False
        Exit Function
    End If
    If Len(txtBuscaEntidad.Text) = 0 Then
        MsgBox "No se ha seleccionado la entidad pagadora al proveedor.", vbInformation, "Aviso"
        txtBuscaEntidad.SetFocus
        ValidaInterfaz = False
        Exit Function
    End If
    If Len(lblCtaDesc.Caption) = 0 Then
        MsgBox "La cuenta de la entidad pagadora al proveedor no existe.", vbInformation
        txtBuscaEntidad.SetFocus
        ValidaInterfaz = False
        Exit Function
    End If
    If fraAbonoSUNAT.Enabled = True Then
        If cboMedioPagoSUNAT.ListIndex = -1 Then
            MsgBox "No se ha seleccionado el tipo de abono a la SUNAT.", vbInformation, "Aviso"
            cboMedioPagoSUNAT.SetFocus
            ValidaInterfaz = False
            Exit Function
        End If
         If Len(txtBuscaEntidad.Text) = 0 Then
            MsgBox "No se ha seleccionado la entidad pagadora a la SUNAT.", vbInformation, "Aviso"
            txtBuscaEntidad.SetFocus
            ValidaInterfaz = False
            Exit Function
        End If
        If Len(lblCtaDesc.Caption) = 0 Then
            MsgBox "La cuenta de la entidad pagadora a SUNAT no existe.", vbInformation
            txtBuscaEntidad.SetFocus
            ValidaInterfaz = False
            Exit Function
        End If
    End If
    ValidaInterfaz = True
End Function
Private Sub cmdGrabar_Click()
Dim lsOpeCod As String
Dim lsOpecod2 As String
Dim lnTipoPago As Integer
Dim row As Integer
Dim lsDocTpo As TpoDoc

Dim oDCapta As DCapMantenimiento
Dim oOpe As DOperacion
Dim oNContFunc  As NContFunciones
Dim oDCtaIF     As DCajaCtasIF
Dim oCtasIF     As NCajaCtaIF
Dim oDocPago      As clsDocPago
Dim oNCaja As nCajaGeneral
Dim oDoc As DDocumento
Set oNCaja = New nCajaGeneral
Dim oImp As NContImprimir
Dim oDis As NRHProcesosCierre

Dim lsdoctipo As String
Dim lsCtaContDebe As String
Dim lsCtaContHaber As String
Dim lsCtaContHaber2 As String
Dim lsCtaEntidadOrig As String
Dim lsTpoIf     As String
Dim lsCtaBanco    As String
Dim lsPersCodIf    As String
Dim lsEntidadOrig    As String
Dim lsSubCuentaIF As String
Dim lsCuentaAho As String
Dim lOk As Integer

Dim lsTpoIFD As String
Dim lsCtaBancoD As String
Dim lsPersCodIFD As String

Dim lsTpoIFH2 As String
Dim lsCtaBancoH2 As String
Dim lsPersCodIFH2 As String

Dim lsDocVoucher As String
Dim lsDocNRo      As String
Dim lsPersCod     As String
Dim lsPersNombre     As String
Dim lnImporte As Currency
Dim lnImporteTotal As Currency
Dim lnImporteITF As Currency 'PASI20150414
Dim lsFecha As String
Dim lsDocumento As String
Dim lsMovNro      As String
Dim lnMovNroPago As Long

Dim lsPlanillaNro As String
Dim lsDocTpoTmp As String
Dim lsDocNroTmp As String, lsDocVoucherTmp As String
Dim lsImpre As String
Dim lsCadBol As String
Dim lsok As String

Dim lsCtaEntidadOrigH2 As String
Dim lsDocTpoH2 As TpoDoc
Dim lsDocNRoH2      As String
Dim lsDocVoucherH2 As String
Dim lsDocTpoTmpH2 As String
Dim lsDocNroTmpH2 As String, lsDocVoucherTmpH2 As String
Dim lsEntidadOrigH2 As String
Dim lsSubCuentaIFH2 As String
Dim lsPlanillaNroH2 As String

    lsOpeCod = gsOpeCod
    lnTipoPago = 0
    lsdoctipo = "-1"
    
    If fnTpoPago = gPagoCuentaCMAC Then
        lsDocTpo = TpoDocNotaAbono
    ElseIf fnTpoPago = gPagoTransferencia Then
        lsDocTpo = TpoDocCarta
    ElseIf fnTpoPago = gPagoCheque Then
        lsDocTpo = TpoDocCheque
    End If
    
    If Trim(Right(cboMedioPagoSUNAT.Text, 50)) = gPagoTransferencia Then
        lsDocTpoH2 = TpoDocCarta
    ElseIf Trim(Right(cboMedioPagoSUNAT.Text, 40)) = gPagoCheque Then
        lsDocTpoH2 = TpoDocCheque
    End If
    
    On Error GoTo NoGrabo
    If Not ValidaInterfaz Then Exit Sub
    row = (fg.row)
    Set oOpe = New DOperacion
    Set oDCapta = New DCapMantenimiento
    If lsDocTpo = TpoDocNotaAbono Then
        If Not oDCapta.CuentaEsValida(fg.TextMatrix(row, 11), fg.TextMatrix(row, 5)) Then
            MsgBox "La cuenta del proveedor " & fg.TextMatrix(row, 5) & " no es correcta o no esta vigente, coordine con el Dpto. de Logistica", vbInformation, "Aviso"
            Set oDCapta = Nothing
        Exit Sub
        End If
    End If
    If lsOpeCod = "" Then
       MsgBox "No se asignó Documentos de Referencia a Operación de Pago", vbInformation, "Aviso"
       Exit Sub
    End If
    lsOpeCod = oOpe.EmiteOpeDoc(Mid(gsOpeCod, 1, 4), lsDocTpo)
    lsOpecod2 = oOpe.EmiteOpeDoc(Mid(gsOpeCod, 1, 4), lsDocTpoH2)
    lsCtaContDebe = fg.TextMatrix(row, 17)
    lsCtaContHaber = oOpe.EmiteOpeCta(lsOpeCod, "H", , txtBuscaEntidad.Text, ObjEntidadesFinancieras)
    lsCtaContHaber2 = IIf(fraAbonoSUNAT.Enabled = True, oOpe.EmiteOpeCta(lsOpecod2, "H", , txtBuscaEntidad2.Text, ObjEntidadesFinancieras), "")
    If lsCtaContDebe = "" Or lsCtaContHaber = "" Then
        MsgBox "Cuentas Contables no determinadas Correctamente" & oImpresora.gPrnSaltoLinea & "consulte con Sistemas", vbInformation, "Aviso"
        Exit Sub
    End If
    Set oDCapta = Nothing
    
    Set oNContFunc = New NContFunciones
    Set oDCtaIF = New DCajaCtasIF
    Set oCtasIF = New NCajaCtaIF
    Set oDocPago = New clsDocPago
    Set oDoc = New DDocumento
    Set oDis = New NRHProcesosCierre
    
    lsCtaEntidadOrig = Trim(lblCtaDesc)
    lsTpoIf = IIf(fnTpoPago = gPagoCuentaCMAC, "", Mid(txtBuscaEntidad, 1, 2))
    lsCtaBanco = IIf(fnTpoPago = gPagoCuentaCMAC, "", Mid(txtBuscaEntidad, 18, Len(Me.txtBuscaEntidad)))
    lsPersCodIf = IIf(fnTpoPago = gPagoCuentaCMAC, "", Mid(txtBuscaEntidad, 4, 13))
    lsEntidadOrig = IIf(fnTpoPago = gPagoCuentaCMAC, "", oDCtaIF.NombreIF(lsPersCodIf))
    lsSubCuentaIF = IIf(fnTpoPago = gPagoCuentaCMAC, "", oCtasIF.SubCuentaIF(lsPersCodIf))
    lsFecha = Format(gdFecSis, "dd/mm/yyyy")
    
    lsCtaEntidadOrigH2 = Trim(lblCtaDesc2)
    lsTpoIFH2 = IIf(Trim(Right(cboMedioPagoSUNAT, 2)) = gPagoCuentaCMAC, "", Mid(txtBuscaEntidad2, 1, 2))
    lsCtaBancoH2 = IIf(Trim(Right(cboMedioPagoSUNAT, 2)) = gPagoCuentaCMAC, "", Mid(txtBuscaEntidad2, 18, Len(Me.txtBuscaEntidad2)))
    lsPersCodIFH2 = IIf(Trim(Right(cboMedioPagoSUNAT, 2)) = gPagoCuentaCMAC, "", Mid(txtBuscaEntidad2, 4, 13))
    lsEntidadOrigH2 = IIf(Trim(Right(cboMedioPagoSUNAT, 2)) = gPagoCuentaCMAC, "", oDCtaIF.NombreIF(lsPersCodIFH2))
    lsSubCuentaIFH2 = IIf(Trim(Right(cboMedioPagoSUNAT, 2)) = gPagoCuentaCMAC, "", oCtasIF.SubCuentaIF(lsPersCodIFH2))
    
    lsTpoIFD = fg.TextMatrix(row, 14)
    lsCtaBancoD = fg.TextMatrix(row, 15)
    lsPersCodIFD = fg.TextMatrix(row, 13)
    
    lsDocVoucher = ""
    lsDocNRo = ""
    lsPersCod = ""
    lnImporte = 0
    
    lsPersCod = fg.TextMatrix(fg.row, 5)
    lsPersNombre = fg.TextMatrix(fg.row, 6)
    lsCuentaAho = fg.TextMatrix(row, 11)
    lnMovNroPago = fg.TextMatrix(row, 9)
    lnImporte = fg.TextMatrix(row, 4)
    lnImporteTotal = fg.TextMatrix(row, 18)
    If lsDocTpo = TpoDocCheque Then
        lsDocVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1))
        oDocPago.InicioCheque lsDocNRo, True, Mid(Me.txtBuscaEntidad, 4, 13), gsOpeCod, lsPersNombre, gsOpeDesc, txtMovDesc.Text, lnImporte, gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, True, gsCodAge, , , lsTpoIf, lsPersCodIf, lsCtaBanco
        If oDocPago.vbOk Then
            lsFecha = oDocPago.vdFechaDoc
            lsDocNroTmp = oDocPago.vsNroDoc
            lsDocVoucherTmp = oDocPago.vsNroVoucher
        Else
            Exit Sub
        End If
    ElseIf lsDocTpo = TpoDocCarta Then
        Do While True
            lsPlanillaNro = InputBox("Ingrese el Nro. de Planilla de Devolución de Pago Proveedor.", "Planilla de Pagos", lsPlanillaNro)
            If lsPlanillaNro = "" Then Exit Sub
            lsPlanillaNro = Format(lsPlanillaNro, "00000000")
            If oDoc.GetValidaDocProv("", CLng(lsDocTpo), lsPlanillaNro) Then
                MsgBox "Nro. de carta ya ha sido ingresada, verifique..!", vbInformation, "Aviso"
            Else
                lsDocNroTmp = lsPlanillaNro
                lsDocVoucherTmp = ""
                gnMgIzq = 17
                gnMgDer = 0
                gnMgSup = 12
                Exit Do
            End If
        Loop
    End If
    If lsDocTpo = TpoDocNotaAbono Then
        lsDocNroTmp = oNContFunc.GeneraDocNro(lsDocTpo, , , , True)
        lsCadBol = oDis.ImprimeBoletaCad(CDate(lsFecha), "ABONO CAJA GENERAL", "Depósito CAJA GENERAL*Nro." & lsDocNRo, "", lnImporte, lsPersNombre, lsCuentaAho, "", 0, 0, "Nota Abono", 0, 0, False, False, , , , True, , , , False, gsNomAge)
    End If
    lsDocNRo = lsDocNroTmp
    If lnImporte < lnImporteTotal Then
        If lsDocTpoH2 = TpoDocCheque Then
            lsDocVoucherH2 = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1))
            oDocPago.InicioCheque lsDocNRoH2, True, Mid(Me.txtBuscaEntidad2, 4, 13), gsOpeCod, "SUNAT", gsOpeDesc, txtMovDesc.Text, (lnImporteTotal - lnImporte), gdFecSis, gsNomCmacRUC, lsSubCuentaIFH2, lsEntidadOrigH2, lsCtaEntidadOrigH2, lsDocVoucherH2, True, gsCodAge, , , lsTpoIFH2, lsPersCodIFH2, lsCtaBancoH2
            If oDocPago.vbOk Then
                lsFecha = oDocPago.vdFechaDoc
                lsDocNroTmpH2 = oDocPago.vsNroDoc
                lsDocVoucherTmpH2 = oDocPago.vsNroVoucher
            Else
                Exit Sub
            End If
        ElseIf lsDocTpoH2 = TpoDocCarta Then
            Do While True
                lsPlanillaNroH2 = InputBox("Ingrese el Nro. de Planilla de Devolución de Pago SUNAT.", "Planilla de Pagos", lsPlanillaNroH2)
                If lsPlanillaNroH2 = "" Then Exit Sub
                lsPlanillaNroH2 = Format(lsPlanillaNroH2, "00000000")
                If oDoc.GetValidaDocProv("", CLng(lsDocTpoH2), lsPlanillaNroH2) Then
                    MsgBox "Nro. de carta ya ha sido ingresada, verifique..!", vbInformation, "Aviso"
                Else
                    lsDocNroTmpH2 = lsPlanillaNroH2
                    lsDocVoucherTmpH2 = ""
                    gnMgIzq = 17
                    gnMgDer = 0
                    gnMgSup = 12
                    Exit Do
                End If
            Loop
        End If
        lsDocNRoH2 = lsDocNroTmpH2
    End If
      Set oDoc = Nothing
    If MsgBox("Desea Grabar la Información", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        cmdGrabar.Enabled = False
        If Not lsDocTpo = TpoDocNotaAbono Then
            lsok = oNCaja.GrabaDevolucionPagoProvSUNAT(gdFecSis, gsCodAge, gsCodUser, gsOpeCod, Trim(Replace(Replace(txtMovDesc.Text, Chr(10), ""), Chr(13), "")), lsCtaContDebe, lsCtaContHaber, lsCtaContHaber2, lsCtaCajaGeneral, lnImporte, lnImporteTotal, _
                    lsPersCod, lsTpoIFD, lsPersCodIFD, lsCtaBancoD, lsTpoIf, lsPersCodIf, lsCtaBanco, lsTpoIFH2, lsPersCodIFH2, lsCtaBancoH2, CStr(lsDocTpo), lsDocNRo, Format(CDate(lsFecha), gsFormatoFecha), lsDocVoucher, _
                    txtNResolucion.Text, Format(CDate(txtFechaDev), gsFormatoFecha), CStr(lsDocTpoH2), lsDocNRoH2, Format(CDate(lsFecha), gsFormatoFecha), lsDocVoucherH2, lsCuentaAho, fnTpoPago, True, lsCtaITFD, lsCtaITFH, gnImpITF, IIf(fnTpoPago = gPagoCuentaCMAC, True, False), _
                    lnMovNroPago, fg.TextMatrix(row, 7))
        Else
            lnImporteITF = oNCaja.DameMontoITF(lnImporte)
            lsMovNro = oNContFunc.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
            lsok = oNCaja.GrabaDevolucionPagoProvSUNATxAbonoCta(lsMovNro, gsOpeCod, Trim(Replace(Replace(txtMovDesc.Text, Chr(10), ""), Chr(13), "")), lsCtaContDebe, lsCtaContHaber, lsCtaContHaber2, lsCtaCajaGeneral, lnImporte, lnImporteTotal, _
                    lsPersCod, lsTpoIFD, lsPersCodIFD, lsCtaBancoD, lsTpoIf, lsPersCodIf, lsCtaBanco, lsTpoIFH2, lsPersCodIFH2, lsCtaBancoH2, CStr(lsDocTpo), lsDocNRo, Format(CDate(lsFecha), gsFormatoFecha), lsDocVoucher, _
                    txtNResolucion.Text, Format(CDate(txtFechaDev), gsFormatoFecha), CStr(lsDocTpoH2), lsDocNRoH2, Format(CDate(lsFecha), gsFormatoFecha), lsDocVoucherH2, lsCuentaAho, fnTpoPago, True, lsCtaITFD, lsCtaITFH, lnImporteITF, IIf(fnTpoPago = gPagoCuentaCMAC, True, False), _
                    lnMovNroPago, fg.TextMatrix(row, 7))
        End If
    Set oNCaja = Nothing
    Set oImp = New NContImprimir
    
    EnviaPrevio lsok, "DEVOLUCIÓN DE RETENCIÓN SUNAT", gnLinPage, False
    
        If lsDocTpo = TpoDocNotaAbono Then
            lsImpre = ""
            lsImpre = lsImpre & lsCadBol
            EnviaPrevio lsImpre & oImpresora.gPrnSaltoPagina, "DEVOLUCIÓN DE RETENCIÓN SUNAT", gnLinPage, False
        End If
        MsgBox "Devolución de Retención se realizó con éxito.", vbInformation, "Mensaje"
    End If
    CargaDatosDevolucionSUNAT
    Set oDCapta = Nothing
    Set oOpe = Nothing
    Set oNContFunc = Nothing
    Set oDCtaIF = Nothing
    Set oCtasIF = Nothing
    Set oDocPago = Nothing
    Set oNCaja = Nothing
    Set oDoc = Nothing
    Set oImp = Nothing
    Set oDis = Nothing
    
    EstablecerDatosInicio
    DesHabilitaControles 0
    cmdGrabar.Enabled = True
    Exit Sub
NoGrabo:
    cmdGrabar.Enabled = True
    Screen.MousePointer = 0
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub
'PASI20150414
'Private Function DameMontoITF(ByVal pnMonto As Double) As Double
'    Dim oITF As DITF
'    Dim lnMonto As Double, lnITF As Double, lnRedondeoITF As Double
'
'    Set oITF = New DITF
'    oITF.fgITFParametros
'
'    lnMonto = pnMonto
'    lnITF = lnMonto * oITF.gnITFPorcent
'    lnITF = oITF.CortaDosITF(lnITF)
'    lnRedondeoITF = oITF.fgDiferenciaRedondeoITF(lnITF)
'    If lnRedondeoITF > 0 Then
'        lnITF = lnITF - lnRedondeoITF
'    End If
'    DameMontoITF = lnITF
'    Set oITF = Nothing
'End Function
'END PASI
Private Sub cmdProcesar_Click()
    CargaDatosDevolucionSUNAT
End Sub
Private Sub cmdRegDevolucion_Click()
    Dim row As Integer
    row = fg.row
    If fg.TextMatrix(row, 1) <> "" Then
        EstablecerDatosProveedorDevolucion (row)
    Else
        MsgBox "No existen Devoluciones Pendientes", vbInformation, "Aviso"
        cmdProcesar.SetFocus
        Exit Sub
    End If
    fg.Enabled = False
    cmdRegDevolucion.Enabled = False
    cmdProcesar.Enabled = False
End Sub
Private Sub EstablecerDatosProveedorDevolucion(ByVal row As Integer)
    Dim oOpe As New DOperacion
    Dim oDCaja As New DCajaGeneral
    Dim rsIfi As ADODB.Recordset
    Set oOpe = New DOperacion
    Set rsIfi = New ADODB.Recordset
    Set oDCaja = New DCajaGeneral
    txtMovDesc.Text = "DEVOLUCION : " & fg.TextMatrix(row, 7)
    If fg.TextMatrix(row, 12) = fsPersCodCMACMaynas Then
        fnTpoPago = gPagoCuentaCMAC
    ElseIf fg.TextMatrix(row, 12) = fsPersCodBCP Then
        fnTpoPago = gPagoTransferencia
    Else
        fnTpoPago = gPagoCheque
    End If
    Select Case fnTpoPago
        Case gPagoCuentaCMAC
            lblMedioPag.Caption = "DEPÓSITO EN CUENTA."
            txtBuscaEntidad.Text = "CAJA MAYNAS"
            lblCtaDesc.Caption = fg.TextMatrix(row, 11)
            If fg.TextMatrix(row, 4) < fg.TextMatrix(row, 18) Then
                fraAbonoSUNAT.Enabled = True
                txtBuscaEntidad2.rs = oOpe.GetRsOpeObj(gsOpeCod, "1")
            Else
                txtBuscaEntidad2.Enabled = False
            End If
        Case gPagoTransferencia
            lblMedioPag.Caption = "TRANSFERENCIA."
            txtBuscaEntidad.rs = oOpe.GetRsOpeObj(gsOpeCod, "1", , , , fsPersCodBCP)
            If fg.TextMatrix(row, 4) < fg.TextMatrix(row, 18) Then
                fraAbonoSUNAT.Enabled = True
                  txtBuscaEntidad2.rs = oOpe.GetRsOpeObj(gsOpeCod, "1")
            Else
                fraAbonoSUNAT.Enabled = False
            End If
        Case gPagoCheque
            lblMedioPag.Caption = "CHEQUE."
            txtBuscaEntidad.rs = oOpe.GetRsOpeObj(gsOpeCod, "1")
             If fg.TextMatrix(row, 4) < fg.TextMatrix(row, 18) Then
                fraAbonoSUNAT.Enabled = True
                txtBuscaEntidad2.rs = oOpe.GetRsOpeObj(gsOpeCod, "1")
            Else
                fraAbonoSUNAT.Enabled = False
            End If
    End Select
    Set rsIfi = Nothing
    DesHabilitaControles fnTpoPago
    txtNResolucion.SetFocus
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub fg_KeyPress(KeyAscii As Integer)
    Dim row As Integer
    row = fg.row
    If KeyAscii = 13 Then
        If fg.TextMatrix(row, 1) <> "" And cmdRegDevolucion.Enabled = True Then
            cmdRegDevolucion.SetFocus
        End If
    End If
End Sub
Private Sub Form_Load()
    Dim oOpe As New DOperacion
    Set oOpe = New DOperacion
    Me.txtFechaDel = DateAdd("m", -1, CDate(gdFecSis))
    Me.txtFechaAl = gdFecSis
    Me.txtFechaDev = gdFecSis
    Me.Caption = "Devolución PAGO SUNAT"
    fsPersCodCMACMaynas = "1090100012521"
    fsPersCodBCP = "1090100824640"
    lsCtaCajaGeneral = "29180703"
    gsOpeCodAsiento2 = "421125"
    Set objPista = New COMManejador.Pista
    lMN = IIf(Mid(gsOpeCod, 3, 1) = Moneda.gMonedaExtranjera, False, True)
    lsFileCarta = App.path & "\" & gsDirPlantillas & gsOpeCod & ".TXT"
    lsCtaITFD = oOpe.EmiteOpeCta(gsOpeCod, "D", 2)
    lsCtaITFH = oOpe.EmiteOpeCta(gsOpeCod, "H", 2)
    EstablecerDatosInicio
    DesHabilitaControles 0
End Sub
Private Sub DesHabilitaControles(ByVal pnEstado As Integer)
    Select Case pnEstado
        Case gPagoCuentaCMAC
            fraDevolucion.Enabled = True
            fraAbono.Enabled = False
        Case gPagoTransferencia, gPagoCheque
            fraDevolucion.Enabled = True
            fraAbono.Enabled = True
        Case Else
            fraDevolucion.Enabled = False
            fraAbono.Enabled = False
            fraAbonoSUNAT.Enabled = False
    End Select
End Sub

Private Sub txtBuscaEntidad_EmiteDatos()
    Dim oCtaIf As NCajaCtaIF
    Set oCtaIf = New NCajaCtaIF
    lblCtaDesc = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaEntidad, 18, 10)) + " " + txtBuscaEntidad.psDescripcion
    If txtBuscaEntidad <> "" Then
        If Not fraAbonoSUNAT.Enabled Then
            Me.cmdGrabar.SetFocus
        Else
            cboMedioPagoSUNAT.SetFocus
        End If
    End If
    Set oCtaIf = Nothing
End Sub
Private Sub txtBuscaEntidad2_EmiteDatos()
     Dim oCtaIf As NCajaCtaIF
    Set oCtaIf = New NCajaCtaIF
    lblCtaDesc2 = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaEntidad2, 18, 10)) + " " + txtBuscaEntidad2.psDescripcion
    If txtBuscaEntidad2 <> "" Then
       Me.cmdGrabar.SetFocus
    End If
    Set oCtaIf = Nothing
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not cmdProcesar.Enabled Then Exit Sub
        Me.cmdProcesar.SetFocus
    End If
End Sub
Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFechaAl.SetFocus
    End If
End Sub
Public Sub CargaDatosDevolucionSUNAT()
    Dim oDCaja As New DCajaGeneral
    Dim row As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    fg.Clear
    fg.Rows = 2
    fg.FormaCabecera
    
    If Not IsDate(txtFechaDel.Text) Then
        MsgBox "La Fecha Desde no es válida. Verifique.", vbInformation, "Aviso"
        Exit Sub
    End If
    If Not IsDate(txtFechaAl.Text) Then
         MsgBox "La Fecha Hasta no es válida. Verifique.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set rs = oDCaja.GetDatosPagoDevolucionProveedorSunat(txtFechaDel, DateAdd("d", 1, CDate(txtFechaAl)))
    Set oDCaja = Nothing
    If rs.EOF Then
        RSClose rs
        MsgBox "No existen Devoluciones Pendientes", vbInformation, "Aviso"
        cmdProcesar.SetFocus
        Exit Sub
    End If
    Do While Not rs.EOF
        fg.AdicionaFila
        fg.ColumnasAEditar = "X-X-X-X-5-X-X-X-X-X-X-X-X-X-X-X-X-X"
        row = fg.row
        fg.TextMatrix(row, 1) = rs!FechaDoc
        fg.TextMatrix(row, 2) = rs!Doc
        fg.TextMatrix(row, 3) = rs!nroDoc
        fg.TextMatrix(row, 4) = Format(rs!nImporteSoles, "#,#0.00")
        fg.TextMatrix(row, 5) = rs!cPersCod
        fg.TextMatrix(row, 6) = rs!Proveedor
        fg.TextMatrix(row, 7) = rs!Glosa
        fg.TextMatrix(row, 8) = rs!cMovNro
        fg.TextMatrix(row, 9) = rs!nMovNro
        fg.TextMatrix(row, 10) = rs!Moneda
        fg.TextMatrix(row, 11) = rs!cCtaCod
        fg.TextMatrix(row, 12) = rs!cPersBanco
        fg.TextMatrix(row, 13) = rs!cPersCodIFI
        fg.TextMatrix(row, 14) = rs!cIFTpo
        fg.TextMatrix(row, 15) = rs!cCtaIfCod
        fg.TextMatrix(row, 16) = rs!nMovNroProv
        fg.TextMatrix(row, 17) = rs!CtaCont
        fg.TextMatrix(row, 18) = Format(rs!nImporteSolesTotal, "#,#0.00")
        rs.MoveNext
    Loop
    RSClose rs
    fg.SetFocus
End Sub
Private Sub txtFechaDev_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If fnTpoPago = gPagoCuentaCMAC Then
            cmdGrabar.SetFocus
        Else
            txtBuscaEntidad.SetFocus
        End If
    End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        txtNResolucion.SetFocus
    End If
End Sub
Private Sub txtNResolucion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFechaDev.SetFocus
    End If
End Sub

