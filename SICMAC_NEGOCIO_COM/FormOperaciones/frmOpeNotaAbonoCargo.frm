VERSION 5.00
Begin VB.Form frmOpeNotaAbonoCargo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   Icon            =   "frmOpeNotaAbonoCargo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCuenta 
      Caption         =   "Cuenta"
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
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton cmdBuscarCuenta 
         Caption         =   "..."
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
   Begin VB.Frame fraCuerpo 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   5055
      Begin VB.ComboBox cboConcepto 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtGlosa 
         Height          =   300
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   3735
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   375
         Left            =   1185
         TabIndex        =   3
         Top             =   1300
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         ForeColor       =   12582912
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Monto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   555
         TabIndex        =   11
         Top             =   1350
         Width           =   600
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3210
         TabIndex        =   10
         Top             =   1380
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto:"
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
         Left            =   270
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Glosa:"
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
         Left            =   600
         TabIndex        =   8
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
End
Attribute VB_Name = "frmOpeNotaAbonoCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************
'** DESARROLLADO POR: FRHU
'** FECHA: 10/02/2015
'** REQUERIMIENTO: ERS048-2014
'*****************************
Option Explicit
Dim nOperacion As COMDConstantes.CaptacOperacion
Dim sMovNroAut As String
Dim nMovVistoElec As Long
Dim sPersCodTit As String
Public Sub Inicio(ByVal nOpe As CaptacOperacion)
    nOperacion = nOpe
    If nOperacion = gCapNotaDeCargo Then
        Me.Caption = "Otras Operaciones - Nota de Cargo"
    Else
        Me.Caption = "Otras Operaciones - Nota de Abono"
    End If
    'FRHU 20150306 OBSERVACION
    txtCuenta.Prod = Trim(Str(gCapAhorros))
    txtCuenta.CMAC = gsCodCMAC
    txtCuenta.EnabledProd = False
    txtCuenta.EnabledCMAC = False
    'FIN FRHU
    Me.Show 1
End Sub
Private Sub Form_Load()
    Call CargarConceptos
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = False
    cboConcepto.ListIndex = 0
    txtCuenta.CMAC = gsCodCMAC
End Sub
Private Sub CargarConceptos()
    Dim oConcepto As New COMDConstantes.DCOMConstantes
    Dim rsConcepto As ADODB.Recordset
    
    Set rsConcepto = oConcepto.GetConceptoNotaCargoAbono(nOperacion)
    Do While Not rsConcepto.EOF
        cboConcepto.AddItem UCase(rsConcepto("cDescripcion")) & space(100) & rsConcepto("nPrdConceptoCod")
        rsConcepto.MoveNext
    Loop
    Set rsConcepto = Nothing
    Set oConcepto = Nothing
End Sub
Public Function VistoElectronico() As Boolean
    Dim loVistoElectronico As New frmVistoElectronico
    Dim lbVistoVal As Boolean
    sMovNroAut = ""
    lbVistoVal = loVistoElectronico.Inicio(12, nOperacion)
                       
    If Not lbVistoVal Then
        MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones.", vbInformation, "Mensaje del Sistema"
        VistoElectronico = False
        Exit Function
    End If
    VistoElectronico = True
    Call loVistoElectronico.RegistraVistoElectronico(0, nMovVistoElec)
End Function
Private Sub cmdAceptar_Click()
    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII
    
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsGen As New COMDConstSistema.DCOMGeneral
    Dim sMovNro, sMovNroTransf As String
    Dim nSaldo As Double, nMonto As Double, nMontoAutoriza As Double
    Dim lsBoleta As String, sCuenta As String
    Dim lsmensaje As String, lsBoletaITF As String
    Dim nConcepto As Integer
    Dim bRechazado As Boolean
    Dim TC As Currency
    
    bRechazado = False
    nMontoAutoriza = CDbl(clsGen.LeeConstSistema(84))
    nConcepto = CInt(Trim(Right(cboConcepto.Text, 3)))
    sCuenta = txtCuenta.NroCuenta
    nMonto = txtMonto.value
    If nMonto = 0 Then
        MsgBox "Debe ingresar un monto diferente de cero"
        txtMonto.SetFocus
        Exit Sub
    End If
    If txtGlosa.Text = "" Then
        MsgBox "Debe ingresar una Glosa"
        txtGlosa.SetFocus
        Exit Sub
    End If
    
    If Not VistoElectronico Then Exit Sub
    
    TC = ObtenerTipoCambio(gdFecSis)
    If IIf(Mid(sCuenta, 9, 1) = "2", nMonto * TC, nMonto) > nMontoAutoriza Then
        If VerificarAutorizacion(bRechazado) = False Then Exit Sub
    End If
    Set clsMov = New COMNContabilidad.NCOMContFunciones 'NContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    
    If MsgBox("Desea grabar la operación?", vbYesNo) = vbNo Then Exit Sub
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    If nOperacion = gCapNotaDeAbono Then
        'nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, , , , , , , , , , sLpt, , , , , gsCodCMAC, , gsCodAge, , gbITFAplica, Me.lblITF.Caption, gbITFAsumidoAho, gITFCobroCargo, , , lsmensaje, lsBoleta, lsBoletaITF, , , , , , , , , gnMovNro, , , , , , , , , , Trim(Right(cboTipoNotaAbono.Text, 5)))
        nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, , , , , , , , , gsNomAge, sLpt, , , , , , gsCodCMAC, , , , False, , , , , lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, , , , , , gnMovNro, , , , , , , , , , , , , , , , , , , nConcepto, nMovVistoElec)
    Else
        nSaldo = clsCap.CapCargoCuentaAho(sCuenta, nMonto, nOperacion, sMovNro, txtGlosa.Text, , , , , , , , , gsNomAge, sLpt, , , , , gsCodCMAC, , , , , , False, , , , lsmensaje, lsBoleta, lsBoletaITF, , , gbImpTMU, , , , , , gnMovNro, , , , , , , , , , nConcepto, nMovVistoElec)
    End If
    If Trim(lsmensaje) <> "" Then MsgBox lsmensaje, vbInformation
    If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
    If Trim(lsBoletaITF) <> "" Then ImprimeBoleta lsBoletaITF, "Boleta ITF"
    Call cmdCancelar_Click
End Sub
Private Sub cmdBuscarCuenta_Click()
    Dim loPers As COMDPersona.UCOMPersona 'UPersona
    Dim lsPersCod As String, lsPersNombre As String
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim lrCtasAhorros As New ADODB.Recordset
    Dim loCuenta As COMDPersona.UCOMProdPersona
    Dim nmoneda As Integer
    
On Error GoTo ControlError

    Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
    Set loPers = Nothing
        
    If Trim(lsPersCod) <> "" Then
        Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
            Set lrCtasAhorros = clsCap.GetProductosAhorros(lsPersCod)
        Set clsCap = Nothing
    End If

    Set loCuenta = New COMDPersona.UCOMProdPersona
    Set loCuenta = frmProdPersona.Inicio(lsPersNombre, lrCtasAhorros)
        If loCuenta.sCtaCod <> "" Then
            txtCuenta.NroCuenta = loCuenta.sCtaCod
            fraCuerpo.Enabled = True
            cmdAceptar.Enabled = True
            cmdCancelar.Enabled = True
            fraCuenta.Enabled = False
            cboConcepto.SetFocus
        Else
            MsgBox "No eligio ninguna cuenta existente o la persona no tiene cuenta", vbInformation
            Exit Sub
        End If
    Set loCuenta = Nothing
    
    nmoneda = CLng(Mid(txtCuenta.NroCuenta, 9, 1))
        
    If nmoneda = gMonedaNacional Then
        txtMonto.BackColor = &HC0FFFF
        lblMon.Caption = "S/."
    Else
        txtMonto.BackColor = &HC0FFC0
        lblMon.Caption = "$"
    End If
    sPersCodTit = lsPersCod
Exit Sub

ControlError:   ' Rutina de control de errores.
MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
       " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
Private Sub ImprimeBoleta(ByVal sBoleta As String, Optional ByVal sMensaje As String = "Boleta Operación")
    Dim nFicSal As Integer
    Do
        nFicSal = FreeFile
        Open sLpt For Output As nFicSal
        Print #nFicSal, sBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        'Print #nFicSal, ""
        Close #nFicSal
    Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub
Private Sub cmdCancelar_Click()
    txtCuenta.Age = ""
    txtCuenta.Prod = ""
    txtCuenta.Cuenta = ""
    txtGlosa.Text = ""
    cboConcepto.ListIndex = 0
    txtMonto.value = 0
    fraCuerpo.Enabled = False
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = False
    fraCuenta.Enabled = True
    'FRHU 20150306 OBSERVACION
    'txtCuenta.SetFocusAge
    txtCuenta.Prod = Trim(Str(gCapAhorros))
    txtCuenta.CMAC = gsCodCMAC
    txtCuenta.EnabledProd = False
    txtCuenta.EnabledCMAC = False
    txtCuenta.SetFocusAge
    'FIN FRHU
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCta As String
        sCta = txtCuenta.NroCuenta
        Call ExisteCuenta(sCta)
        
        '***marg ers065-2017***
        Dim nmoneda As Integer
        nmoneda = CLng(Mid(txtCuenta.NroCuenta, 9, 1))
        
        If nmoneda = gMonedaNacional Then
            txtMonto.BackColor = &HC0FFFF
            lblMon.Caption = "S/."
        Else
            txtMonto.BackColor = &HC0FFC0
            lblMon.Caption = "$"
        End If
        'end marg***************
        
    End If
End Sub
Private Sub ObtenerTitularCuenta()
    Dim oTitular As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rs As ADODB.Recordset
    'Set rs = oTitular.ObtenerTitularCuenta(txtCuenta.NroCuenta)
    'Set rs = oTitular.GetPersonaCuenta(txtCuenta.NroCuenta, gCapRelPersRepTitular) 'FRHU 20150708 INCIDENTE 'COMENTADO BY GEMO20200805
    Set rs = oTitular.GetPersonaCuenta(txtCuenta.NroCuenta, gCapRelPersTitular)  'COMENTADO BY GEMO20200805
    If Not rs.BOF And Not rs.EOF Then
        'sPersCodTit = rs!codigo
        sPersCodTit = rs!cPersCod 'FRHU 20150708 INCIDENTE
    End If
    Set oTitular = Nothing
    Set rs = Nothing
End Sub
Private Sub ExisteCuenta(ByVal psCuenta As String)
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sMsg As String
    
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    sMsg = clsCap.ValidaCuentaOperacion(psCuenta, True)
    Set clsCap = Nothing
    If sMsg = "" Then
        fraCuerpo.Enabled = True
        cmdAceptar.Enabled = True
        cmdCancelar.Enabled = True
        Call ObtenerTitularCuenta
    Else
        MsgBox sMsg, vbInformation, "Operacion"
        txtCuenta.SetFocusAge
    End If
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtMonto.SetFocus 'txtGlosa.SetFocus 'FRHU 20150306 OBSERVACION
End Sub
Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdAceptar.SetFocus
End Sub
Private Function VerificarAutorizacion(ByRef pbRechazado As Boolean) As Boolean

Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
Dim oCapAutN  As COMNCaptaGenerales.NCOMCaptAutorizacion
Dim rs As New ADODB.Recordset

Dim lsmensaje As String
Dim nMonto As Double
Dim cMoneda As String
'Dim lbRechazado As Boolean

nMonto = txtMonto.value
cMoneda = Mid(txtCuenta.NroCuenta, 9, 1)
   
Set oCapAutN = New COMNCaptaGenerales.NCOMCaptAutorizacion
If sMovNroAut = "" Then 'Si es nueva, registra nueva solicitud
    
    oCapAutN.NuevaSolicitudOtrasOperaciones sPersCodTit, IIf(nOperacion = gCapNotaDeCargo, "2", "3"), gdFecSis, nMonto, cMoneda, Trim(txtGlosa.Text), gsCodUser, gOpeAutorizacionNotaCargoAbono, gsCodAge, sMovNroAut, nMovVistoElec
        
    Do While VerificarAutorizacion = False
        If Not oCapAutN.VerificarAutorizacionOtrasOperaciones(IIf(nOperacion = gCapNotaDeCargo, "2", "3"), nMonto, sMovNroAut, lsmensaje, pbRechazado) Then
            If lsmensaje = "Esta Operación Aun no esta Autorizada" Then
                If MsgBox("Para proceder con la operacion debe solicitar VºBº del Jefe de Operaciones..." & vbNewLine & _
                          "Desea continuar esperando la Autorización?", vbYesNo) = vbNo Then
                    Exit Do
                Else
                    VerificarAutorizacion = False
                End If
            End If
            If lsmensaje = "Esta Operación fue Rechazada" Then
                MsgBox lsmensaje, vbInformation, "Aviso"
                Exit Do
            End If
        Else
            MsgBox lsmensaje, vbInformation, "Aviso"
            VerificarAutorizacion = True
        End If
    Loop
    
 End If
Set oCapAutN = Nothing
End Function
Private Function ObtenerTipoCambio(ByVal pdFechaSiniestro As Date) As Currency
    Dim oGen As New COMDConstSistema.DCOMGeneral
    ObtenerTipoCambio = oGen.GetTipCambio(pdFechaSiniestro, 7)
End Function
