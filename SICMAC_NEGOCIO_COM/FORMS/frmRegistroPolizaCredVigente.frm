VERSION 5.00
Begin VB.Form frmRegistroPolizaCredVigente 
   Caption         =   "Mantenimiento Gastos Póliza contraincendio"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8460
   Icon            =   "frmRegistroPolizaCredVigente.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   6240
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   33
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6480
      TabIndex        =   32
      Top             =   5040
      Width           =   855
   End
   Begin VB.ComboBox cboHasta 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   5160
      Width           =   855
   End
   Begin VB.Frame Frame6 
      Caption         =   "Aplicación en cuotas"
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   4920
      Width           =   6255
      Begin VB.TextBox txtMontoCuota 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cboDesde 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Monto Prima Cuota:"
         Height          =   255
         Left            =   3360
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   1800
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Desde:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Prima"
      Height          =   1455
      Left            =   3960
      TabIndex        =   16
      Top             =   3360
      Width           =   4335
      Begin VB.TextBox txtTC 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtPrimaNeta 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtPrimaTotalTC 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtPrimaTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblPrimaTotalTC 
         Caption         =   "Prima total T/C:"
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "T/C Cierre mes:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Prima Neta:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Prima total:"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos del Seguro"
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   3735
      Begin VB.TextBox txtMontoValorEdific 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtClaseInmueble 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   14
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtMonedaValorEdific 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtNumCertificado 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Clase Inmueble:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Valor de Edificación:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Nº de Certificado:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Aseguradora"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   8175
      Begin SICMACT.TxtBuscar txtAseguradora 
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
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
         Enabled         =   0   'False
         Enabled         =   0   'False
         EnabledText     =   0   'False
      End
      Begin VB.Label txtNombreAseg 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3480
         TabIndex        =   35
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Aseguradora:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame frmRegistroPolizaCredVigente 
      Caption         =   "Garantía"
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   8175
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "Aplicar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6840
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin SICMACT.FlexEdit grdGarantia 
         Height          =   975
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   1720
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Garantía-Código"
         EncabezadosAnchos=   "300-4500-1500"
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
         ColumnasAEditar =   "X-X-X"
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         Texto           =   "Crédito:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmRegistroPolizaCredVigente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnInmueble As Integer
Dim R As ADODB.Recordset

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObtieneDatosCuenta (AXCodCta.NroCuenta)
    End If
End Sub

Private Sub cboDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 0 Then
        KeyAscii = 0
        'MsgBox "No permitido", vbCritical, "Aviso"
    End If
End Sub


Private Sub cboHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 0 Then
        KeyAscii = 0
        'MsgBox "No permitido", vbCritical, "Aviso"
    End If
End Sub

Private Sub cmdAplicar_Click()
    Dim loCred As COMDCredito.DCOMCreditos
    Dim oPol As COMDCredito.DCOMPoliza
    Dim rsGarant As ADODB.Recordset
    
    
    Dim rs As ADODB.Recordset
    Set loCred = New COMDCredito.DCOMCreditos
    Set rsGarant = New ADODB.Recordset
    
    Set oPol = New COMDCredito.DCOMPoliza
    
    Set rsGarant = loCred.RecuperaValorEdificacion(grdGarantia.TextMatrix(grdGarantia.row, 2))
    
    
    Call CargarDatosSeguro(rsGarant)
    
    'Call frmCredPolizaListado.Inicio(0)
    'Set rs = oPol.CargaDatosPoliza(frmCredPolizaListado.sNumPoliza)
    
    
    'Dim oPol As COMDCredito.DCOMPoliza
    'Dim nEstPol As Integer
    'Call frmCredPolizaListado.Inicio(nBusqueda)
    'Set rs = oPol.CargaDatosPoliza(frmCredPolizaListado.sNumPoliza)
End Sub

Private Sub cmdBuscar_Click()
    Dim loPers As COMDPersona.UCOMPersona 'UPersona
    Dim lsPersCod As String, lsPersNombre As String
    Dim lsEstados As String
    'Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
    Dim loPersCredito  As COMDCredito.DCOMCreditos
    
    Dim lrCreditos As New ADODB.Recordset
    Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

    Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
    Set loPers = Nothing
    

    If Trim(lsPersCod) <> "" Then
        Set loPersCredito = New COMDCredito.DCOMCreditos
            Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod)
        Set loPersCredito = Nothing
    End If

    Set loCuentas = New COMDPersona.UCOMProdPersona
        Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
        If loCuentas.sCtaCod <> "" Then
            AXCodCta.Enabled = True
            AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
            AXCodCta.SetFocusCuenta
        End If
    Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
    Dim loCred As COMDCredito.DCOMCredito
    Dim loCreditos As COMDCredito.DCOMCreditos
    Dim rsCta As ADODB.Recordset
    Dim rsGarant As ADODB.Recordset
    Dim rsCuotaCred As ADODB.Recordset
    
    Dim nCta As String
    Dim nITF As Double
    
    Dim i As Integer
    
    Dim nNumFilas As Integer
    
    Dim lnCuotaIni As Integer
    Dim lnCuotaFin As Integer
    
    Set loCred = New COMDCredito.DCOMCredito
    Set loCreditos = New COMDCredito.DCOMCreditos
    Set rsCta = New ADODB.Recordset
    Set rsGarant = New ADODB.Recordset
    Set rsCuotaCred = New ADODB.Recordset
    
    'set rsCta = oColRec.ObtenerPagoCredAdjudicacion(sCuenta, "136303")
    Set rsGarant = loCred.RecuperaGarantiasCredito(sCuenta)
    Set rsCuotaCred = loCreditos.ObtenerUltimaCuotaPag(sCuenta)
    Set R = loCred.RecuperaDatosCreditoVigente(sCuenta, , True)
    'grdGarantia.AdicionaFila
    If Not (rsGarant.EOF And rsGarant.BOF) Then
        grdGarantia.FormaCabecera
        nNumFilas = rsGarant.RecordCount
        For i = 1 To nNumFilas
            grdGarantia.TextMatrix(i, 0) = i
            grdGarantia.TextMatrix(i, 1) = rsGarant!cDescripcion
            grdGarantia.TextMatrix(i, 2) = rsGarant!cNumGarant
            If i <> nNumFilas Then
                grdGarantia.AdicionaFila
            End If
            rsGarant.MoveNext
        Next
    Else
        MsgBox "No se ha encontrado información de la cuenta ingresada"
        AXCodCta.SetFocus
    End If
    
    
    
    If Not (rsCuotaCred.EOF And rsCuotaCred.BOF) Then
        lnCuotaIni = rsCuotaCred!nCuota
        lnCuotaFin = rsCuotaCred!nCuotaUlt
        If lnCuotaIni > lnCuotaFin Then
            MsgBox "Crédito tiene una condición no valida"
            Exit Sub
        End If
        Call CargarComboCuotas(lnCuotaIni, lnCuotaFin)
    Else
        MsgBox "No se ha encontrado información de las Cuotas del Crèdito ingresado"
        Exit Sub
    End If
    
    cmdAplicar.Enabled = True
    cmdCancelar.Enabled = True
    cmdBuscar.Enabled = False
    AXCodCta.Enabled = False
End Sub

Public Sub CargarDatosSeguro(ByVal poDRGarant As ADODB.Recordset)
    Dim lnPrimaNeta As Double
    Dim lnPrimaTotal As Double
    Dim lnPrimatotalTC As Double
    Dim lnValorEdificacion As Double
    Dim lnNtna As Double
    Dim lnPrimaNetaMin As Double
    Dim loCred As COMDCredito.DCOMCreditos
    Dim rsCotSeguro As ADODB.Recordset
    Dim lnMontoxIGV As Double
    Dim lnDerechoEmision As Double
    
    Set loCred = New COMDCredito.DCOMCreditos
    Set rsCotSeguro = New ADODB.Recordset
    
    gnIGVValor = CargaImpuestoFechaValor(gcCtaIGV, gdFecSis) / 100
    
    If Not (poDRGarant.EOF And poDRGarant.BOF) Then
        txtMonedaValorEdific.Text = poDRGarant!cMoneda
        lnValorEdificacion = poDRGarant!nValorEdificacion
        txtMontoValorEdific.Text = Format(lnValorEdificacion, gsFormatoNumeroView)
        txtClaseInmueble.Text = poDRGarant!cConsDescripcion
        lnInmueble = poDRGarant!nInmueble
        GetTipCambioLog gdFecSis, False
        txtTC.Text = Format(gnTipCambio, "#.00")
    Else
        MsgBox "No se ha encontrado información de la garantia ingresada", vbCritical, "Aviso"
        Exit Sub
    End If
    If txtMonedaValorEdific.Text = "S/." Then
        lnValorEdificacion = lnValorEdificacion / gnTipCambio
        lblPrimaTotalTC.Visible = True
        txtPrimaTotalTC.Visible = True
    End If
    Set rsCotSeguro = loCred.ObtenerCotizacionSeguroXAgencia(gsCodAge, lnInmueble)
    
    If Not (rsCotSeguro.EOF And rsCotSeguro.BOF) Then
        lnNtna = rsCotSeguro!ntna
        lnPrimaNetaMin = rsCotSeguro!nPrimaNetaMinima
        
        lnPrimaNeta = (lnValorEdificacion * lnNtna) / 1000
        If lnPrimaNeta < lnPrimaNetaMin Then
            lnPrimaNeta = lnPrimaNetaMin
        End If
        
        lnDerechoEmision = (lnPrimaNeta * rsCotSeguro!ndrchoEmision) / 100
        lnPrimaNeta = lnPrimaNeta + lnDerechoEmision
        txtPrimaNeta.Text = Format(lnPrimaNeta, "#.00")
    Else
        MsgBox "No se ha encontrado información de la garantia ingresada"
        Exit Sub
    End If
    
    lnMontoxIGV = lnPrimaNeta * gnIGVValor
    lnPrimaTotal = lnPrimaNeta + lnMontoxIGV
    txtPrimaTotal.Text = Format(lnPrimaTotal, "#.00")
    If txtMonedaValorEdific.Text <> "S/." Then
        txtPrimaTotalTC.Text = Format(lnPrimaTotal, "#.00")
        txtMontoCuota.Text = Format(lnPrimaTotal / 12, "#.00")
    Else
        txtPrimaTotalTC.Text = Format(lnPrimaTotal * gnTipCambio, "#.00")
        txtMontoCuota.Text = Format(CDbl(txtPrimaTotalTC.Text) / 12, "#.00")
    End If
    
    
    cmdGrabar.Enabled = True
    txtAseguradora.Enabled = True
End Sub

Public Sub CargarComboCuotas(ByVal nCuotaIni As Integer, ByVal nCuotaFin As Integer)
    Dim X As Integer
    For X = nCuotaIni To nCuotaFin
        cboDesde.AddItem (X)
        'cboHasta.AddItem (X)
    Next
    cboHasta.AddItem (nCuotaFin)
    cboDesde.ListIndex = 0
    cboHasta.ListIndex = 0
End Sub

Private Sub cmdCancelar_Click()
     Call LimpiarFormulario
End Sub

Private Sub CmdGrabar_Click()
    Dim oBase As COMDCredito.DCOMCredActBD
    Dim oGasto As COMDCredito.DCOMGasto
    Dim oCred As COMDCredito.DCOMCreditos
    Dim lnCuotaDesde As Integer
    Dim lnCuotaHasta As Integer
    Dim lnNroCuota As Integer
    Dim oDR As New ADODB.Recordset
    Set oGasto = New COMDCredito.DCOMGasto
    Set oCred = New COMDCredito.DCOMCreditos
    
    Set oDR = New ADODB.Recordset
    Set oBase = New COMDCredito.DCOMCredActBD
    
    
    If AXCodCta.Age <> gsCodAge Then
        MsgBox "No puede realizar la operación debido a que el crédito pertenece a otra agencia", vbCritical, "Aviso"
        Exit Sub
    End If
    If txtNumCertificado.Text <> "" And txtAseguradora.Text <> "" Then
        
        
        If oBase.ValidaNumeroCertificado(txtAseguradora.Text, txtNumCertificado.Text) = True Then
            MsgBox "No puede realizar la operación debido a que el Numero de Certificado ya existe para la aseguradora", vbCritical, "Aviso"
            txtNumCertificado.SetFocus
            Exit Sub
        End If
        Dim MatCuotas() As Integer
        Dim i As Integer
        Dim J As Integer
        Dim sMensaje As String
        
        lnCuotaDesde = CDbl(cboDesde.Text)
        lnCuotaHasta = CDbl(cboHasta.Text)
        
        If lnCuotaDesde > lnCuotaHasta Then
            MsgBox "La cuota 'Desde' no puede ser mayor que la cuota 'Hasta'", vbCritical, "Aviso"
            Exit Sub
        End If
        For i = lnCuotaDesde To lnCuotaHasta - 1
            lnNroCuota = lnNroCuota + 1
        Next
        
        ReDim MatCuotas(lnNroCuota + 1)
        For i = lnCuotaDesde To lnCuotaHasta
            MatCuotas(J) = i
            J = J + 1
        Next
        
       
        sMensaje = oBase.InsercionGastosxCuotaLote(AXCodCta.NroCuenta, R!nNroCalen, 1, MatCuotas, 1231, CDbl(txtMontoCuota.Text), True)
        oBase.MantenimientoGastoPolizaContraincendio AXCodCta.NroCuenta, gsCodUser, gsCodAge, Format(gdFecSis, "yyyy/MM/dd"), _
        CDbl(cboDesde.Text), CDbl(cboHasta.Text), CDbl(txtPrimaNeta.Text), CDbl(txtPrimaTotal.Text), _
        CDbl(txtPrimaTotalTC.Text), CDbl(txtMontoCuota.Text), txtNumCertificado.Text, txtAseguradora.Text
        MsgBox "Los datos se guardaron correctamente", vbExclamation, "Aviso"
        LimpiarFormulario
        'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Inserta Gasto: " & CLng(Trim(Right(CboGasto.Text, 15))) & ", Cuota: Todas", ActxCuenta.NroCuenta, gCodigoCuenta
                
        Set oBase = Nothing
        Set oDR = Nothing
        If sMensaje <> "" Then
            MsgBox sMensaje, vbInformation, "Mensaje"
        End If
    
    Else
        If txtNumCertificado.Text = "" Then
            txtNumCertificado.SetFocus
            MsgBox "Ingrese el Nº de Certificado", vbCritical, "Aviso"
            Exit Sub
        End If
        If txtAseguradora.Text = "" Then
            txtAseguradora.SetFocus
            MsgBox "Ingrese el Aseguradora", vbCritical, "Aviso"
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
   Call LimpiarFormulario
End Sub
Public Sub LimpiarFormulario()
    AXCodCta.NroCuenta = ""
    grdGarantia.Clear
    txtMonedaValorEdific.Text = ""
    txtMontoValorEdific.Text = "0.00"
    txtClaseInmueble.Text = ""
    txtTC.Text = "0.00"
    txtPrimaNeta.Text = "0.00"
    txtPrimaTotal.Text = "0.00"
    txtPrimaTotalTC.Text = "0.00"
     cboDesde.Clear
    cboHasta.Clear
    txtMontoCuota.Text = "0.00"
    cmdAplicar.Enabled = False
    cmdBuscar.Enabled = True
    AXCodCta.Enabled = True
    cmdGrabar.Enabled = False
    txtAseguradora.Enabled = False
    cmdCancelar.Enabled = False
    lblPrimaTotalTC.Visible = False
    txtPrimaTotalTC.Visible = False
    txtNombreAseg.Caption = ""
    txtAseguradora.Text = ""
    txtNumCertificado.Text = ""
End Sub

Private Sub txtAseguradora_EmiteDatos()
    Dim loPers As COMDPersona.UCOMPersona 'UPersona
    Dim lsPersCod As String, lsPersNombre As String
    
    Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
    Set loPers = Nothing
    
    txtAseguradora.Text = lsPersCod
    txtNombreAseg.Caption = lsPersNombre
End Sub
