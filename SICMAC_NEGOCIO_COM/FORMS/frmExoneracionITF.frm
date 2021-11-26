VERSION 5.00
Begin VB.Form frmExoneracionITF 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6675
      Left            =   45
      TabIndex        =   3
      Top             =   90
      Width           =   8100
      Begin VB.Frame FraHistorico 
         Caption         =   "Histórico de Exoneraciones"
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
         Height          =   2790
         Left            =   90
         TabIndex        =   17
         Top             =   3735
         Width           =   7935
         Begin SICMACT.FlexEdit grdHistorico 
            Height          =   2445
            Left            =   105
            TabIndex        =   18
            Top             =   270
            Width           =   7680
            _ExtentX        =   13547
            _ExtentY        =   4313
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            VisiblePopMenu  =   -1  'True
            EncabezadosNombres=   "#-Fecha-Tipo Exoneraciòn-CodExo"
            EncabezadosAnchos=   "250-1800-5500-0"
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X"
            ListaControles  =   "0-0-0-0"
            EncabezadosAlineacion=   "C-C-L-C"
            FormatosEdit    =   "0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            ColWidth0       =   255
            RowHeight0      =   300
            TipoBusPersona  =   1
         End
      End
      Begin VB.Frame fraCliente 
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
         Height          =   2475
         Left            =   60
         TabIndex        =   6
         Top             =   615
         Width           =   7950
         Begin SICMACT.FlexEdit grdCliente 
            Height          =   1755
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   7680
            _ExtentX        =   13811
            _ExtentY        =   3096
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            VisiblePopMenu  =   -1  'True
            EncabezadosNombres=   "#-Codigo-Nombre-Relacion-Direccion-ID"
            EncabezadosAnchos=   "250-1700-3800-1500-0-0"
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-L-L-C-C"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            ColWidth0       =   255
            RowHeight0      =   300
            TipoBusPersona  =   1
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cuenta :"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   2138
            Width           =   960
         End
         Begin VB.Label lblTipoCuenta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   1155
            TabIndex        =   12
            Top             =   2085
            Width           =   1440
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "# Firmas :"
            Height          =   195
            Left            =   2640
            TabIndex        =   11
            Top             =   2145
            Width           =   690
         End
         Begin VB.Label lblFirmas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   3465
            TabIndex        =   10
            Top             =   2085
            Width           =   465
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "TipoMoneda: "
            Height          =   195
            Left            =   4440
            TabIndex        =   9
            Top             =   2160
            Width           =   990
         End
         Begin VB.Label lblMoneda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   5520
            TabIndex        =   8
            Top             =   2085
            Width           =   705
         End
      End
      Begin VB.CheckBox chkExo 
         Caption         =   "Exoneraciòn"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1845
         TabIndex        =   5
         Top             =   3135
         Width           =   1395
      End
      Begin VB.ComboBox cboTipoExoneracion 
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3375
         Width           =   7860
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   105
         TabIndex        =   14
         Top             =   270
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Exoneraciòn"
         Height          =   195
         Left            =   75
         TabIndex        =   16
         Top             =   3150
         Width           =   1470
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   4125
         TabIndex        =   15
         Top             =   300
         Width           =   3960
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   45
      TabIndex        =   2
      Top             =   6795
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5940
      TabIndex        =   1
      Top             =   6840
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7065
      TabIndex        =   0
      Top             =   6840
      Width           =   1000
   End
End
Attribute VB_Name = "frmExoneracionITF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By capi 21012009
Option Explicit
Dim objPista As COMManejador.Pista
'End by

Private Sub CargaTpoExoneracion()
Dim ssql As String, i As Integer
Dim rstemp As ADODB.Recordset
Dim miconst As COMDConstantes.DCOMConstantes

    Set miconst = New COMDConstantes.DCOMConstantes
    Set rstemp = miconst.GetConstante("1044", True, True, True, , "0','1044")
    Set miconst = Nothing
    i = 0
    cboTipoExoneracion.Clear
    While Not rstemp.EOF
        cboTipoExoneracion.AddItem rstemp!Columna1
        cboTipoExoneracion.ItemData(i) = rstemp!Columna2
        i = i + 1
        rstemp.MoveNext
    Wend
End Sub

Private Sub CHKeXO_Click()
    If chkExo.value = 1 Then
        cboTipoExoneracion.Enabled = True
    Else
        cboTipoExoneracion.Enabled = False
    End If
    cboTipoExoneracion.ListIndex = -1
End Sub

Private Sub cmdCancelar_Click()
    LimpiaControles
End Sub

Private Sub LimpiaControles()
grdCliente.Clear
grdCliente.rows = 2
grdCliente.FormaCabecera
grdHistorico.Clear
grdHistorico.rows = 2
grdHistorico.FormaCabecera

lblmoneda.Caption = "S/."
lblmoneda.BackColor = &HC0FFFF
lblMensaje = ""
CmdGrabar.Enabled = False
cmdcancelar.Enabled = False
lblFirmas = ""
lblTipoCuenta = ""
fraCliente.Enabled = False
FraHistorico.Enabled = False
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Prod = ""
txtCuenta.Cuenta = ""
txtCuenta.EnabledCMAC = False
txtCuenta.EnabledAge = False
txtCuenta.EnabledProd = True
If cboTipoExoneracion.ListCount > 0 Then
    cboTipoExoneracion.ListIndex = -1
End If
chkExo.value = vbUnchecked
cboTipoExoneracion.Enabled = False
End Sub


Private Sub cmdGrabar_Click()
    Dim ssql As String, Aux As String, TpoExo As String
    Dim VCMovNro As String
    Dim oMov As COMDMov.DCOMMov
    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim sValor As String
                    
On Error GoTo Mensaje
    'FRHU 20171102: Acta 189-2017
    If Me.chkExo.value = 1 And Me.cboTipoExoneracion.Text = "" Then
        MsgBox "Debe Elegir un tipo de exoneracion Valido.", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Se va a Grabar los Datos, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Sleep (500)
    'FIN FRHU 20171102
    Set oMov = New COMDMov.DCOMMov
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    VCMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    If cboTipoExoneracion.ListIndex = -1 Then
        TpoExo = "0"
    Else
        TpoExo = cboTipoExoneracion.ItemData(cboTipoExoneracion.ListIndex)
    End If
    
    'FRHU 20171102 : Acta 189-2017
    'If Me.chkExo.value = 1 And Me.cboTipoExoneracion.Text = "" Then
        'MsgBox "Debe Elegir un tipo de exoneracion Valido.", vbInformation, "Aviso"
        'Exit Sub
    'End If
    'FIN FRHU 20171102
    'Set oConecta = New DConecta
      If oCap.TieneExoneracion(txtCuenta.NroCuenta) = 0 Then
         sValor = IIf(chkExo.value = vbChecked, TpoExo, "0")
         oCap.InsertarITFExoneracion txtCuenta.NroCuenta, sValor, VCMovNro
      Else
         sValor = IIf(chkExo.value = vbChecked, TpoExo, "0")
         Aux = IIf(chkExo.value = vbUnchecked, ", cExtorno='" & VCMovNro & "'", "")
         oCap.ActualizarITFExoneracion txtCuenta.NroCuenta, sValor, VCMovNro, Aux
      End If
      sValor = IIf(chkExo.value = vbChecked, TpoExo, "0")
      oCap.InsertarITFExoneracionDET txtCuenta.NroCuenta, sValor, VCMovNro
      
    'By Capi 21012009
     objPista.InsertarPista gsOpeCod, VCMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, , txtCuenta.NroCuenta, gCodigoCuenta
    'End by
            

      
      If cboTipoExoneracion.ListCount > 0 Then
         cboTipoExoneracion.ListIndex = -1
      End If
      chkExo.value = vbUnchecked
      ObtieneHistoricoExo txtCuenta.NroCuenta
       
       Exit Sub
       Set oMov = Nothing
       Set oCap = Nothing
Mensaje:
        MsgBox "Ocurrio un error: " & err.Number & " " & err.Description, vbOKOnly + vbExclamation, "Error"
       
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim nProducto As Producto
 If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(nProducto, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
    'By Capi 20012009
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCapExoneraITF
    'End By


End Sub

Private Sub Form_Load()
CargaTpoExoneracion
cboTipoExoneracion.Enabled = False
txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledCMAC = False
txtCuenta.EnabledAge = True
Me.Caption = "Exoneración de ITF"
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCta As String
    sCta = txtCuenta.NroCuenta
    LimpiaControles
    txtCuenta.Prod = Mid(sCta, 6, 3)
    txtCuenta.Cuenta = Right(sCta, 10)
    
    ObtieneDatosCuenta sCta
    ObtieneHistoricoExo sCta
End If
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim rsCta As ADODB.Recordset, rsRel As ADODB.Recordset
Dim nEstado As COMDConstantes.CaptacEstado
Dim nRow As Long
Dim sMsg As String, sMoneda As String, sPersona As String
Dim nDocumento As Integer, nPersoneria As Integer, nmoneda As Integer, nProducto As Integer
Dim sTipoCuenta As String
Dim bDocumento As Boolean
Dim nTipoCuenta As Integer
Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
sMsg = clsCap.ValidaCuentaOperacion(sCuenta)
Set clsCap = Nothing
If sMsg = "" Then
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        If bDocumento And nDocumento = TpoDocOrdenPago Then
            If Not rsCta("bOrdPag") Then
                rsCta.Close
                Set rsCta = Nothing
                MsgBox "Cuenta NO fue aperturada con ORDEN DE PAGO", vbInformation, "Aviso"
                txtCuenta.Cuenta = ""
                txtCuenta.SetFocus
                Exit Sub
            End If
        End If
        nEstado = rsCta("nPrdEstado")
        nPersoneria = rsCta("nPersoneria")
        'lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy hh:mm:ss")
        nmoneda = CLng(Mid(sCuenta, 9, 1))
        If nmoneda = gMonedaNacional Then
            sMoneda = "MONEDA NACIONAL"
            'txtMonto.BackColor = &HC0FFFF
            lblmoneda.BackColor = &HC0FFFF
            lblmoneda.Caption = "S/."
        Else
            sMoneda = "MONEDA EXTRANJERA"
            'txtMonto.BackColor = &HC0FFC0
            lblmoneda.BackColor = &HC0FFC0
            lblmoneda.Caption = "US$"
        End If
        Select Case nProducto
            Case gCapAhorros
                If rsCta("bOrdPag") Then
                    lblMensaje = "AHORROS CON ORDEN DE PAGO" & Chr$(13) & sMoneda
                Else
                    lblMensaje = "AHORROS SIN ORDEN DE PAGO" & Chr$(13) & sMoneda
                End If
                'lblUltContacto = Format$(rsCta("dUltContacto"), "dd mmm yyyy hh:mm:ss")
            Case gCapPlazoFijo
                'lblUltContacto = rsCta("nPlazo")
            Case gCapCTS
                'lblUltContacto = rsCta("cInstitucion")
        End Select
        lblTipoCuenta = UCase(rsCta("cTipoCuenta"))
        sTipoCuenta = lblTipoCuenta
        nTipoCuenta = rsCta("nPrdCtaTpo")
        lblFirmas = Format$(rsCta("nFirmas"), "#0")
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        sPersona = ""
        Do While Not rsRel.EOF
            If rsRel("cPersCod") = gsCodPersUser Then
                MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                Unload Me
                Exit Sub
            End If
            If sPersona <> rsRel("cPersCod") Then
                grdCliente.AdicionaFila
                nRow = grdCliente.rows - 1
                grdCliente.TextMatrix(nRow, 1) = rsRel("cPersCod")
                grdCliente.TextMatrix(nRow, 2) = UCase(PstaNombre(rsRel("Nombre")))
                grdCliente.TextMatrix(nRow, 3) = UCase(rsRel("Relacion")) & Space(50) & Trim(rsRel("nPrdPersRelac"))
                grdCliente.TextMatrix(nRow, 4) = rsRel("Direccion")
                grdCliente.TextMatrix(nRow, 5) = rsRel("ID N°")
                sPersona = rsRel("cPersCod")
            End If
            rsRel.MoveNext
        Loop
        
        
        rsRel.Close
        Set rsRel = Nothing
        Set clsMant = Nothing
        Set clsCap = Nothing
        fraCliente.Enabled = True
        FraHistorico.Enabled = True
    
        CmdGrabar.Enabled = True
       
        cmdcancelar.Enabled = True
    End If
    
   ' MuestraFirmas sCuenta
    
Else
    MsgBox sMsg, vbInformation, "Operacion"
    txtCuenta.SetFocus
End If
End Sub

Private Sub ObtieneHistoricoExo(ByVal sCuenta As String)
Dim ssql As String, i As Integer
Dim rstemp As ADODB.Recordset
Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
Set rstemp = New ADODB.Recordset
Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rstemp = oCap.ObtenerHistorialExonearcion(sCuenta)
i = 1
If Not rstemp Is Nothing Then
While Not rstemp.EOF
    grdHistorico.AdicionaFila
    grdHistorico.TextMatrix(i, 1) = Right(rstemp!cMovNro, 2) & "/" & Mid(rstemp!cMovNro, 5, 2) & "/" & Left(rstemp!cMovNro, 4)
    grdHistorico.TextMatrix(i, 2) = rstemp!cConsDescripcion
    grdHistorico.TextMatrix(i, 3) = rstemp!Nexotpo
    
    i = i + 1
    rstemp.MoveNext
Wend
End If
Set oCap = Nothing
Set rstemp = Nothing
End Sub

