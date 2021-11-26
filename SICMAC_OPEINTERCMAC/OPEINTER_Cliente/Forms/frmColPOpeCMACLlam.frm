VERSION 5.00
Begin VB.Form frmColPOpeCMACLlam 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   Icon            =   "frmColPOpeCMACLlam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5580
      TabIndex        =   6
      Top             =   2760
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5580
      TabIndex        =   5
      Top             =   2280
      Width           =   990
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5580
      TabIndex        =   7
      Top             =   3240
      Width           =   990
   End
   Begin VB.Frame fraMovimiento 
      Caption         =   "Datos de Tarjeta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2535
      Left            =   60
      TabIndex        =   9
      Top             =   1680
      Width           =   5490
      Begin VB.TextBox txtTrack2 
         Height          =   315
         Left            =   1200
         TabIndex        =   22
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtClave 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   20
         Top             =   650
         Width           =   1095
      End
      Begin VB.TextBox txtTarjeta 
         Height          =   315
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   17
         Top             =   240
         Width           =   3375
      End
      Begin VB.Frame Frame1 
         Caption         =   "Glosa"
         Height          =   975
         Left            =   90
         TabIndex        =   15
         Top             =   1470
         Width           =   2565
         Begin VB.TextBox txtGlosa 
            Height          =   630
            Left            =   60
            TabIndex        =   3
            Top             =   225
            Width           =   2415
         End
      End
      Begin VB.Frame fraMonto 
         Height          =   975
         Left            =   2685
         TabIndex        =   12
         Top             =   1470
         Width           =   2715
         Begin VB.TextBox txtMonto 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1065
            MaxLength       =   14
            TabIndex        =   4
            Top             =   360
            Width           =   1545
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
            Height          =   195
            Left            =   90
            TabIndex        =   16
            Top             =   405
            Width           =   540
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Track2:"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1110
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Clave:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   675
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nº Tarjeta:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   270
         Width           =   765
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Datos Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   6450
      Begin VB.TextBox txtCuenta 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         MaxLength       =   20
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmColPOpeCMACLlam.frx":030A
         Left            =   960
         List            =   "frmColPOpeCMACLlam.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2535
      End
      Begin VB.TextBox txtCliente 
         Height          =   315
         Left            =   960
         MaxLength       =   39
         TabIndex        =   2
         Top             =   1140
         Width           =   4215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   330
         Width           =   585
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00800000&
         Height          =   795
         Left            =   3720
         TabIndex        =   13
         Top             =   240
         Width           =   2475
      End
      Begin VB.Image imagen 
         Height          =   480
         Index           =   0
         Left            =   5880
         Picture         =   "frmColPOpeCMACLlam.frx":030E
         Top             =   1020
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   720
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmColPOpeCMACLlam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************
'* LLAMADADA A CMAC DE OPERACIONE CREDITO PIGNORATICIO
'Archivo:  frmColOpeCMACLlam.frm
'LAYG   :  25/07/2001.
'Resumen:  Nos registrar la llamada a otra CMAC de un a operacion de cred pignoraticio

Option Explicit
Dim nOpeCod As Long
Dim sOpeDesc As String
Dim sPersCodCMAC As String
Dim sNombreCMAC As String
Dim nmoneda As Moneda

Public Sub Inicio(ByVal pnOpeCod As Long, ByVal psOpeDesc As String, _
        ByVal psPersCodCMAC As String, ByVal psNomCmac As String, _
        Optional nComision As Double = 0)

nOpeCod = pnOpeCod
sOpeDesc = psOpeDesc
sPersCodCMAC = psPersCodCMAC
sNombreCMAC = psNomCmac
lblMensaje = sNombreCMAC & Chr$(13) & sOpeDesc

If pnOpeCod = 107001 Then
    Me.Caption = "Créditos CMAC Llamada - " & psOpeDesc
    gsOpeCod = CStr(pnOpeCod)
End If

txtMonto.Text = 0#
txtMonto.Text = Format$(txtMonto, "#,##0.00")
'Select Case fnVarOpeCod
'    Case geColPRenDEOtCj
'        txtDocumento.Visible = True
'End Select
cboMoneda.AddItem "SOLES" & Space(100) & gMonedaNacional
cboMoneda.AddItem "DOLARES" & Space(100) & gMonedaExtranjera
'cboMoneda.ListIndex = 0
'lblComision = Format$(nComision, "#,##0.00")
Me.Show 1
End Sub

Private Sub cboMoneda_Change()
Dim vMonto As Double
Dim vNComision As Double
Dim vTipoCambio As COMDConstSistema.NCOMTipoCambio
Set vTipoCambio = New COMDConstSistema.NCOMTipoCambio

If Val(txtMonto.Text) = 0 Then
    cmdGrabar.Enabled = False
Else
     vMonto = GetLimiteMonto(nmoneda)
   
'    If Val(txtMonto.Text) > vMonto Then
'        vNComision = GetComisionMayor()
'       If nmoneda = gMonedaNacional Then
'        lblComision.Caption = Format(((Val(txtMonto.Text) * vNComision) * 0.01), "#,##0.00")
'       Else
'        lblComision.Caption = Format(((Val(txtMonto.Text) * vNComision) * 0.01) * vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta), "#,##0.00")
'       End If
'    Else
'        vNComision = GetValorComision()
'       lblComision.Caption = Format(vNComision, "#0.00")
'    End If


    cmdGrabar.Enabled = True
End If
Set vTipoCambio = Nothing
End Sub

Private Sub cboMoneda_Click()
nmoneda = Right(cboMoneda.Text, 2)
If nmoneda = gMonedaNacional Then
    txtMonto.BackColor = &HC0FFFF
Else
    txtMonto.BackColor = &HC0FFC0
End If

Dim vMonto As Double
Dim vNComision As Double

'Dim vTipoCambio As COMDConstSistema.NCOMTipoCambio
'Set vTipoCambio = New COMDConstSistema.NCOMTipoCambio

Dim vTipoCambio As DFunciones.dFuncionesNeg
Set vTipoCambio = New DFunciones.dFuncionesNeg

If Val(txtMonto.Text) = 0 Then
    cmdGrabar.Enabled = False
Else
     vMonto = GetLimiteMonto(nmoneda)
   
'    If Val(txtMonto.Text) > vMonto Then
'        vNComision = GetComisionMayor()
'       If nmoneda = gMonedaNacional Then
'        lblComision.Caption = Format(((Val(txtMonto.Text) * vNComision) * 0.01), "#,##0.00")
'       Else
'        lblComision.Caption = Format(((Val(txtMonto.Text) * vNComision) * 0.01) * vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta), "#,##0.00")
'       End If
'    Else
'        vNComision = GetValorComision()
'       lblComision.Caption = Format(vNComision, "#0.00")
'    End If


    cmdGrabar.Enabled = True
End If
Set vTipoCambio = Nothing
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCuenta.SetFocus
End If
End Sub

Private Sub cmdGrabar_Click()
Dim lsCuenta As String
Dim lsClave As String
Dim lnMonto As Currency
Dim lsNumTarj As String
Dim lnMoneda As Integer
Dim lsTrack2 As String
Dim lsmensaje As String
Dim lsBoleta As String
Dim lsBoletaITF As String

'Dim loLavDinero As SICMACMOPEINTER.frmMovLavDinero
'Set loLavDinero = New SICMACMOPEINTER.frmMovLavDinero


lsCuenta = Trim(txtCuenta)
lsNumTarj = Trim(txtTarjeta)
lnMonto = Val(txtMonto.Text)
lsClave = Trim(txtClave)
lsTrack2 = Trim(txtTrack2)

If lsCuenta = "" Then
    MsgBox "Debe digitar un número de cuenta válido", vbInformation, "Aviso"
    txtCuenta.SetFocus
    Exit Sub
End If
If lsNumTarj = "" Then
    MsgBox "Debe digitar un número de tarjeta", vbInformation, "Aviso"
    txtCliente.SetFocus
    Exit Sub
End If
If lnMonto = 0 Then
    MsgBox "Debe colocar un monto mayor a cero", vbInformation, "Aviso"
    txtMonto.SetFocus
    Exit Sub
End If
If lsClave = "" Then
    MsgBox "Debe digitar una clave válida", vbInformation, "Aviso"
    txtCuenta.SetFocus
    Exit Sub
End If
If lsTrack2 = "" Then
    MsgBox "Debe digitar un Track2", vbInformation, "Aviso"
    txtCuenta.SetFocus
    Exit Sub
End If

'On Error GoTo ControlError
'Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim lsFun As DFunciones.dFuncionesNeg
Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lsIFTipo As String
Dim lsGlosa As String
Dim sNumTarj As String

Set lsFun = New DFunciones.dFuncionesNeg
lsIFTipo = Format$(gTpoIFCmac, "00")
lsGlosa = Trim(txtGlosa.Text)

If cboMoneda.ListIndex = 0 Then
    lnMoneda = 1
Else
    lnMoneda = 2
End If

If MsgBox(" Desea Grabar la Operación - " & sOpeDesc & " ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    
    'Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios
    Dim nMontoLavDinero As Double, nTC As Double
    Dim sPersLavDinero As String

    'Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
'    If Not clsExo.EsCuentaExoneradaLavadoDinero(lsCuenta) Or Len(lsCuenta) < 18 Then
'        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
'        Set clsExo = Nothing
'        sPersLavDinero = ""
'        'nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
'        Set clsLav = Nothing
'        If lnMoneda = gMonedaNacional Then
'            Dim clsTC As COMDConstSistema.NCOMTipoCambio
'            Set clsTC = New COMDConstSistema.NCOMTipoCambio
'            nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
'            Set clsTC = Nothing
'        Else
'            nTC = 1
'        End If
'        If lnMonto >= Round(nMontoLavDinero * nTC, 2) Then
'            'By Capi 1402208
'            Call IniciaLavDinero(loLavDinero)
'            'ALPA 20081009****************************************************
'            'sperslavdinero = loLavDinero.Inicia(, , , , False, True, CDbl(txtMonto.Text), lsCuenta, Mid(Me.Caption, 15), False, "", , , , , lnMoneda)
'            sPersLavDinero = loLavDinero.Inicia(, , , , False, True, CDbl(txtMonto.Text), lsCuenta, Mid(Me.Caption, 15), False, "", , , , , lnMoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
'            '*****************************************************************
'            If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
'            'End
'
'            'sperslavdinero = IniciaLavDinero()
'            'If sperslavdinero = "" Then Exit Sub
'        End If
'    Else
'        Set clsExo = Nothing
'    End If
    
    
    'Dim oCap As COMNCaptaGenerales.NCOMCaptaGenerales
    'Dim sCuentaAho As String
    'Set oCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    'sCuentaAho = lsFun.GetCuentaAbonoIF(sPersCodCMAC, nmoneda)
    'lsCuenta = sCuentaAho
    'Set oCap = Nothing
    
'    If sCuentaAho = "" Then
'        MsgBox "No Existe cuenta de regularización.", vbInformation, "Aviso"
'        Exit Sub
'    End If
    
    cmdGrabar.Enabled = False
        
        'Genera el Mov Nro
        'Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = lsFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        'Set loContFunct = Nothing
        Call RegistrarOperacionInterCMAC(lsNumTarj, lsClave, lsCuenta, nOpeCod, lsTrack2, nmoneda, "", sPersCodCMAC, lsMovNro, sLpt, sOpeDesc, sNombreCMAC, lsCuenta, lnMonto, lsGlosa, lsIFTipo, False)

Dim nVez As Integer
nVez = 1
        'Impresión
'        Call ImprimirRecibo(nVez, lsCuenta)
'
'        Do While True
'            If MsgBox("Desea reimprimir boletas de regularizacion?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'                nVez = nVez + 1
'                Call ImprimirRecibo(nVez, lsCuenta)
'            Else
'                Exit Do
'            End If
'        Loop
        'Limpiar
        
        Unload Me
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
 
End Sub

'Private Sub IniciaLavDinero(poLavDinero As SICMACMOPEINTER.frmMovLavDinero)
'
'Dim sPersCod As String
'Dim sNombre As String
'Dim sDireccion As String
'Dim sDocId As String
'Dim nMonto As Double
'Dim sCuenta As String
'Dim n_Moneda As Integer
'
'If cboMoneda.ListIndex = 0 Then
'    n_Moneda = 1
'ElseIf cboMoneda.ListIndex = 1 Then
'    n_Moneda = 2
'End If
'
'    poLavDinero.TitPersLavDinero = "" 'LblCodCli.Caption
'    poLavDinero.TitPersLavDineroNom = "" 'LblNomCli.Caption
'    poLavDinero.TitPersLavDineroDir = "" 'LblCliDirec.Caption
'    poLavDinero.TitPersLavDineroDoc = "" 'LblDocNat.Caption
'
'    'sPersCod = "" 'LblCodCli.Caption
'    'sNombre = "" 'LblNomCli.Caption
'    'sDireccion = "" 'LblCliDirec.Caption
'    'sDocId = "" 'LblDocNat.Caption
'
'
'    sCuenta = txtCuenta.Text
'
'    nMonto = CDbl(txtMonto.Text)
'    'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, True, nMonto, sCuenta, Trim(fnVarOpeCod), True, "COLOCACIONES", , , , , n_Moneda)
'End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras(KeyAscii)
If KeyAscii = 13 Then
'    If Me.txtDocumento.Enabled And Me.txtDocumento.Visible Then
'        txtDocumento.SetFocus
'    Else
        'txtMonto.SetFocus
        'txtGlosa.SetFocus
        txtTarjeta.SetFocus
'    End If
End If
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCliente.SetFocus
    Exit Sub
End If
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.txtGlosa.Enabled And Me.txtGlosa.Visible Then
        txtGlosa.SetFocus
    Else
        txtMonto.SetFocus
    End If
End If
End Sub

'Private Sub txtExtracto_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If txtDocumento.Visible Then
'        txtDocumento.SetFocus
'    Else
'        txtMonto.SetFocus
'    End If
'    Exit Sub
'End If
'KeyAscii = NumerosEnteros(KeyAscii)
'End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    Me.txtMonto.SetFocus
End If
End Sub

Private Sub txtMonto_Change()
'Dim vMonto As Double
'Dim vNComision As Double
'Dim vTipoCambio As COMDConstSistema.NCOMTipoCambio
'Set vTipoCambio = New COMDConstSistema.NCOMTipoCambio
'
'If txtMonto.value = 0 Then
'    cmdGrabar.Enabled = False
'Else
'     vMonto = GetLimiteMonto(fnVarMoneda)
'
'    If txtMonto.value > vMonto Then
'        vNComision = GetComisionMayor()
'       If fnVarMoneda = gMonedaNacional Then
'        lblComision.Caption = Format(((txtMonto.value * vNComision) * 0.01), "#,##0.00")
'       Else
'        lblComision.Caption = Format(((txtMonto.value * vNComision) * 0.01) * vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta), "#,##0.00")
'       End If
'    Else
'        vNComision = GetValorComision()
'       lblComision.Caption = Format(vNComision, "#0.00")
'    End If
'
'
'    cmdGrabar.Enabled = True
'End If
'Set vTipoCambio = Nothing
End Sub
Private Function GetLimiteMonto(ByVal nmoneda As Moneda) As Double

Dim nValor As Double
'Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim clsFun As DFunciones.dFuncionesNeg
'Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
Set clsFun = New DFunciones.dFuncionesNeg
 'nValor = 0
nValor = clsFun.GetCapParametro(IIf(nmoneda = gMonedaNacional, 2078, 2079))
    
Set clsFun = Nothing

GetLimiteMonto = nValor

End Function

Private Function GetComisionMayor() As Double

Dim rsPar As New ADODB.Recordset
'Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim clsFun As DFunciones.dFuncionesNeg
'Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
Set clsFun = New DFunciones.dFuncionesNeg

Set rsPar = clsFun.GetTarifaParametro(nOpeCod, gMonedaNacional, 2077)
Set clsFun = Nothing
If rsPar.EOF And rsPar.BOF Then
    GetComisionMayor = 0
Else
    GetComisionMayor = rsPar("nParValor")
End If
rsPar.Close
Set rsPar = Nothing

End Function

Private Function GetValorComision() As Double
Dim rsPar As New ADODB.Recordset
'Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim clsFun As DFunciones.dFuncionesNeg
'Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
Set clsFun = New DFunciones.dFuncionesNeg

Set rsPar = clsFun.GetTarifaParametro(nOpeCod, gMonedaNacional, gCostoOperacionCMACLlam)
Set clsFun = Nothing
If rsPar.EOF And rsPar.BOF Then
    GetValorComision = 0
Else
    GetValorComision = rsPar("nParValor")
End If
rsPar.Close
Set rsPar = Nothing
End Function

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    Dim vMonto As Double
    Dim vNComision As Double
    'Dim vTipoCambio As COMDConstSistema.NCOMTipoCambio
    If KeyAscii = 13 Then
        'Set vTipoCambio = New COMDConstSistema.NCOMTipoCambio
        If Val(txtMonto.Text) = 0 Then
            cmdGrabar.Enabled = False
        Else
             vMonto = GetLimiteMonto(nmoneda)
             'vMonto = 999999#
           
'            If Val(txtMonto.Text) > vMonto Then
'                vNComision = GetComisionMayor()
'               If nmoneda = gMonedaNacional Then
'                lblComision.Caption = Format(((Val(txtMonto.Text) * vNComision) * 0.01), "#,##0.00")
'               Else
'                lblComision.Caption = Format(((Val(txtMonto.Text) * vNComision) * 0.01) * vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta), "#,##0.00")
'               End If
'            Else
'               'vNComision = GetValorComision()
'               vNComision = 20#
'               lblComision.Caption = Format(vNComision, "#0.00")
'            End If
        
        
            cmdGrabar.Enabled = True
        End If
        'Set vTipoCambio = Nothing
        If cmdGrabar.Enabled Then cmdGrabar.SetFocus
    End If
End Sub

'Private Sub txtSaldo_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    txtGlosa.SetFocus
'End If
'End Sub

'Private Sub ImprimirRecibo(ByVal nVez As Integer, ByVal psCuenta As String)
'Dim N As COMNCaptaGenerales.NCOMCaptaMovimiento
'Set N = New COMNCaptaGenerales.NCOMCaptaMovimiento
'
'Dim lsBoleta As String
'Dim nFicSal As String
'
'lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'lsBoleta = lsBoleta & N.ImprimeBoletaCMAC("CREDITO CMAC LLAMADA", fsVarOpeDesc, Trim(txtMonto.Text), Trim(txtCliente.Text), Trim(psCuenta), "", 0, fnVarMoneda, "", 0, 0, False, False, , , fsVarNombreCMAC, , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodAge, , , gbImpTMU) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'
'lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'
'lsBoleta = lsBoleta & N.ImprimeBoletaCMAC("CREDITO CMAC LLAMADA", "Comision CMAC Llam", Trim(lblComision.Caption), Trim(txtCliente.Text), Trim(psCuenta), "", 0, gMonedaNacional, "", 0, 0, False, False, , , fsVarNombreCMAC, , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodAge, , , gbImpTMU) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'lsBoleta = lsBoleta & N.ImprimeBoletaCMACRegula("CREDITO CMAC LLAMADA", fsVarOpeDesc, Trim(txtMonto.Text), Trim(txtCliente.Text), Trim(psCuenta), "", 0, fnVarMoneda, "", 0, 0, False, False, , , fsVarNombreCMAC, , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodAge, False, txtCuenta.Text) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'If Trim(lsBoleta) <> "" Then
'      nFicSal = FreeFile
'      Open sLpt For Output As nFicSal
'          Print #nFicSal, lsBoleta
'          Print #nFicSal, ""
'      Close #nFicSal
'End If
'Set N = Nothing
'End Sub
