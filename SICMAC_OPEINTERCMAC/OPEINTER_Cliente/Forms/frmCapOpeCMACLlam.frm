VERSION 5.00
Begin VB.Form frmCapOpeCMACLlam 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "frmCapOpeCMACLlam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   4605
      TabIndex        =   4
      Top             =   4110
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5685
      TabIndex        =   5
      Top             =   4110
      Width           =   1000
   End
   Begin VB.Frame fraMovimiento 
      Caption         =   "Datos Tarjeta"
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
      Height          =   2550
      Left            =   60
      TabIndex        =   7
      Top             =   1515
      Width           =   6675
      Begin VB.TextBox txtTarjeta 
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
         Left            =   1200
         MaxLength       =   22
         TabIndex        =   18
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtClave 
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
         IMEMode         =   3  'DISABLE
         Left            =   5220
         MaxLength       =   4
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   360
         Width           =   1155
      End
      Begin VB.Frame fraGlosa 
         Caption         =   "Glosa"
         Height          =   1485
         Left            =   90
         TabIndex        =   14
         Top             =   960
         Width           =   2925
         Begin VB.TextBox txtGlosa 
            Height          =   1125
            Left            =   75
            TabIndex        =   3
            Top             =   240
            Width           =   2730
         End
      End
      Begin VB.Frame fraMonto 
         Caption         =   "Track2"
         Height          =   1485
         Left            =   3015
         TabIndex        =   10
         Top             =   960
         Width           =   3570
         Begin VB.TextBox txtTrack2 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox txtMonto 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Left            =   960
            MaxLength       =   14
            TabIndex        =   15
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   900
            Width           =   540
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "N° Tarjeta:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   420
         Width           =   765
      End
      Begin VB.Label lblEtqDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Clave :"
         Height          =   195
         Left            =   4620
         TabIndex        =   19
         Top             =   420
         Width           =   495
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
      Height          =   1455
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   6660
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
         MaxLength       =   18
         TabIndex        =   1
         Top             =   660
         Width           =   2535
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmCapOpeCMACLlam.frx":030A
         Left            =   960
         List            =   "frmCapOpeCMACLlam.frx":030C
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
         Top             =   1020
         Width           =   4215
      End
      Begin VB.Label Label6 
         Caption         =   "Moneda :"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   330
         Width           =   675
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
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
         Height          =   675
         Left            =   3720
         TabIndex        =   12
         Top             =   240
         Width           =   2715
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1110
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta :"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   720
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmCapOpeCMACLlam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nOperacion As CaptacOperacion
Dim sPersCodCMAC As String
Dim sNombreCMAC As String
Dim sDescOperacion As String
Dim nmoneda As Moneda
Dim nMontoMinimoComMaynas As Double
Dim nComisionPorcentMaynas As Double
Public Sub Inicia(ByVal nOpe As CaptacOperacion, ByVal sDescOpe As String, _
        ByVal sCodCmac As String, ByVal sNomCmac As String, Optional nComision As Double = 0)

sDescOperacion = sDescOpe
Me.Caption = "Captaciones - Operaciones InterCMACs - " & sDescOperacion
nOperacion = nOpe
sPersCodCMAC = sCodCmac
sNombreCMAC = sNomCmac
lblMensaje = sNombreCMAC & Chr$(13) & sDescOperacion

'Select Case nOperacion
'    Case gCMACOTAhoDepChq
'        lblEtqDocumento.Visible = True
'        txtDocumento.Visible = True
'        lblEtqDocumento = "Cheque N° :"
'    Case gCMACOTAhoRetOP
'        lblEtqDocumento.Visible = True
'        txtDocumento.Visible = True
'        lblEtqDocumento = "Orden Pago N° :"
'    Case Else
'        lblEtqDocumento.Visible = False
'        txtDocumento.Visible = False
'End Select

If nOperacion = gCMACOTAhoDepChq Or nOperacion = gCMACOTAhoDepEfec Then
    txtMonto.ForeColor = &HC00000
ElseIf nOperacion = gCMACOTAhoRetEfec Or nOperacion = gCMACOTAhoRetOP Then
    txtMonto.ForeColor = &HC0&
End If

txtMonto.Text = 0#
txtMonto.Text = Format(txtMonto, "#,##0.00")
gsOpeCod = CStr(nOperacion)

cboMoneda.AddItem "SOLES" & Space(100) & gMonedaNacional
cboMoneda.AddItem "DOLARES" & Space(100) & gMonedaExtranjera
'Me.cboMoneda.ListIndex = 0
Me.Show 1
End Sub

Private Sub cboMoneda_Change()
Dim vMonto As Double
Dim vNComision As Double

'Dim vTipoCambio As COMDConstSistema.NCOMTipoCambio
'Set vTipoCambio = New COMDConstSistema.NCOMTipoCambio
Dim vTipoCambio As DFunciones.dFuncionesNeg
Set vTipoCambio = New DFunciones.dFuncionesNeg

If Val(txtMonto.Text) = 0 Then
    cmdGrabar.Enabled = False
Else
'    If gbITFAplica Then
'        Me.lblITF.Caption = Format(fgITFCalculaImpuesto(CCur(txtMonto.Text)), "#,##0.00")
'    End If
    
    vMonto = GetLimiteMonto(nmoneda)
   
'    If CDbl(Val(txtMonto.Text)) > vMonto Then
'       vNComision = GetComisionMayor()
'       If nmoneda = gMonedaNacional Then
'            lblComision.Caption = Format(((Val(txtMonto.Text) * vNComision) * 0.01), "#,##0.00")
'       Else
'            lblComision.Caption = Format(((Val(txtMonto.Text) * vNComision) * 0.01) * vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta), "#,##0.00")
'       End If
'    Else
'            vNComision = GetValorComision()
'            lblComision.Caption = Format(vNComision, "#0.00")
'    End If

    cmdGrabar.Enabled = True
End If
Set vTipoCambio = Nothing
End Sub

Private Sub cboMoneda_Click()
'*********modificado para que solo cobre comision en SOLES

nmoneda = Right(cboMoneda.Text, 2)
If nmoneda = gMonedaNacional Then
    txtMonto.BackColor = &HC0FFFF
    'txtSaldo.BackColor = &HC0FFFF
Else
    txtMonto.BackColor = &HC0FFC0
    'txtSaldo.BackColor = &HC0FFC0
End If

Dim vMonto As Double
Dim vNComision As Double

'Dim vTipoCambio As COMDConstSistema.NCOMTipoCambio
'Set vTipoCambio = New COMDConstSistema.NCOMTipoCambio

Dim vTipoCambio As DFunciones.dFuncionesNeg
Set vTipoCambio = New DFunciones.dFuncionesNeg

txtMonto.Text = Format(txtMonto, "#,##0.00")
If Val(txtMonto.Text) = 0 Then
    cmdGrabar.Enabled = False
Else
    vMonto = GetLimiteMonto(nmoneda)
'    If CDbl(Val(txtMonto.Text)) > vMonto Then
'       vNComision = GetComisionMayor()
'       If nmoneda = gMonedaNacional Then
'            lblComision.Caption = Format(((Val(txtMonto.Text) * vNComision) * 0.01), "#,##0.00")
'       Else
'            lblComision.Caption = Format(((Val(txtMonto.Text) * vNComision) * 0.01) * vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta), "#,##0.00")
'       End If
'    Else
'        vNComision = GetValorComision()
'        lblComision.Caption = Format(vNComision, "#0.00")
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
Dim sCuenta As String
Dim sClave As String
Dim nMonto As Double
Dim sDocumento As String
Dim nTC As Double
Dim sNumTarj As String
Dim sGlosa As String
Dim sMovNro As String
Dim sCtaAbono As String
Dim sTrack2 As String

Dim clsFun As DFunciones.dFuncionesNeg
'Dim loLavDinero As SICMACMOPEINTER.frmMovLavDinero
'Set loLavDinero = New SICMACMOPEINTER.frmMovLavDinero


sNumTarj = Trim(txtTarjeta.Text)
sCuenta = Trim(txtCuenta.Text)
nMonto = Val(txtMonto.Text)
sClave = Trim(txtClave.Text)
sTrack2 = Trim(txtTrack2.Text)


If sCuenta = "" Then
    MsgBox "Debe digitar un número de cuenta válido", vbInformation, "Aviso"
    txtCuenta.SetFocus
    Exit Sub
End If
If sNumTarj = "" Then
    MsgBox "Debe digitar un número de tarjeta válido", vbInformation, "Aviso"
    txtTarjeta.SetFocus
    Exit Sub
End If
If sTrack2 = "" Then
    MsgBox "Debe digitar el Tack2", vbInformation, "Aviso"
    txtTrack2.SetFocus
    Exit Sub
End If
If nMonto = 0 Then
    MsgBox "Debe colocar un monto mayor a cero", vbInformation, "Aviso"
    txtMonto.SetFocus
    Exit Sub
End If

If MsgBox("Desea Grabar la Operacion??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Set clsFun = New DFunciones.dFuncionesNeg
    If nOperacion = "260501" Or nOperacion = "260503" Then

        'Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios
        Dim nMontoLavDinero As Double
        Dim sPersLavDinero As String, sReaPersLavDinero As String, sBenPersLavDinero As String
        
        'Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
        'If Not clsExo.EsCuentaExoneradaLavadoDinero(sCuenta) Or Len(sCuenta) < 18 Then
'        If Len(sCuenta) < 18 Then
'            'Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
'            'Set clsExo = Nothing
'            sPersLavDinero = ""
'            'nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
'            nMontoLavDinero = clsFun.GetCapParametro(gMonMensLavDineroME)
'            'Set clsLav = Nothing
'            If nmoneda = gMonedaNacional Then
'                'Dim clsTC As COMDConstSistema.NCOMTipoCambio
'                'Set clsTC = New COMDConstSistema.NCOMTipoCambio
'                'nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
'                nTC = clsFun.EmiteTipoCambio(gdFecSis, TCFijoDia)
'                'Set clsTC = Nothing
'            Else
'                nTC = 1
'            End If
'            If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
'                'By Capi 1402208
'                    Call IniciaLavDinero(loLavDinero)
'                    'ALPA 20081009************************************************************
'                    sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), True, "", , , , , nmoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
'                    If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
'                'End
'            End If
'        Else
'            'Set clsExo = Nothing
'        End If

    End If
        
    sGlosa = txtGlosa.Text
     
    sMovNro = clsFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

'    sCtaAbono = clsFun.GetCuentaAbonoIF(sPersCodCMAC, nmoneda)
    
    Dim lsBoleta As String
    Dim nFicSal As Integer
    
'    If Len(Trim(sCtaAbono)) = 0 Then
'        MsgBox "No existe cuenta con la cual hacer la regularización", vbExclamation, "Aviso"
'        Exit Sub
'    End If
    
    Call RegistrarOperacionInterCMAC(sNumTarj, sClave, sCuenta, nOperacion, sTrack2, nmoneda, "", sPersCodCMAC, sMovNro, sLpt, sDescOperacion, sNombreCMAC, sCuenta, nMonto, sGlosa, "", False)
    
    gVarPublicas.LimpiaVarLavDinero
    'Set loLavDinero = Nothing
    Set clsFun = Nothing
    Unload Me
End If
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
'sPersCod = "" 'LblCodCli.Caption
'sNombre = "" 'LblNomCli.Caption
'sDireccion = "" 'LblCliDirec.Caption
'sDocId = "" 'LblDocNat.Caption
'
'poLavDinero.TitPersLavDinero = ""
'poLavDinero.TitPersLavDineroNom = ""
'poLavDinero.TitPersLavDineroDir = ""
'poLavDinero.TitPersLavDineroDoc = ""
'
'
'If cboMoneda.ListIndex = 0 Then
'    n_Moneda = 1
'ElseIf cboMoneda.ListIndex = 1 Then
'    n_Moneda = 2
'End If
'
'    sCuenta = txtCuenta.Text
'
'nMonto = CDbl(txtMonto.Text)
''IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, True, nMonto, sCuenta, Trim(nOperacion), True, "COLOCACIONES", , , , , n_Moneda)
'End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    nMontoMinimoComMaynas = GetMontoMinimoMaynas
    nComisionPorcentMaynas = GetPorcentMaynas
    txtMonto.Text = Format(txtMonto, "#,##0.00")
End Sub

Private Function GetMontoMinimoMaynas() As Double
Dim nValor As Double
Dim clsFun As DFunciones.dFuncionesNeg
Set clsFun = New DFunciones.dFuncionesNeg

nValor = clsFun.GetCapParametro(2048)

Set clsFun = Nothing
GetMontoMinimoMaynas = nValor
End Function
Private Function GetPorcentMaynas() As Double '
Dim nValor As Double
Dim clsFun As DFunciones.dFuncionesNeg
Set clsFun = New DFunciones.dFuncionesNeg

nValor = clsFun.GetCapParametro(2049)

Set clsFun = Nothing
GetPorcentMaynas = nValor
End Function

Private Sub lblComision_Change()
    If gbITFAplica Then
    '****PREGUNTAR SI SE COBRARA ITF POR COMISION
'        Me.lblITF.Caption = Format(fgITFCalculaImpuesto(CCur(lblComision.caption)), "#,##0.00")
    End If
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras(KeyAscii)
If KeyAscii = 13 Then
    txtGlosa.SetFocus
    Exit Sub
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
    txtMonto.SetFocus
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
    txtMonto.SetFocus
    If cmdGrabar.Enabled = True Then
        cmdGrabar.SetFocus
    End If
End If
End Sub

Private Function GetLimiteMonto(ByVal nmoneda As Moneda) As Double

Dim nValor As Double
'Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
'Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
Dim clsFun As New dFuncionesNeg
Set clsFun = New dFuncionesNeg

 nValor = 0
 nValor = clsFun.GetCapParametro(IIf(nmoneda = gMonedaNacional, 2078, 2079))

Set clsFun = Nothing
GetLimiteMonto = nValor
End Function

Private Function GetComisionMayor() As Double

Dim rsPar As New ADODB.Recordset
'Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
'Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
Dim clsFun As New dFuncionesVarias
Set clsFun = New dFuncionesVarias
    Set rsPar = clsFun.GetTarifaParametro(nOperacion, gMonedaNacional, 2077)
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
'Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
Dim clsFun As New dFuncionesVarias
Set clsFun = New dFuncionesVarias

Set rsPar = clsFun.GetTarifaParametro(nOperacion, gMonedaNacional, gCostoOperacionCMACLlam)
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
   Dim vTipoCambio As DFunciones.dFuncionesNeg

   KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 15, 2)
   If KeyAscii = 13 Then
        'Set vTipoCambio = New COMDConstSistema.NCOMTipoCambio
        Set vTipoCambio = New DFunciones.dFuncionesNeg
        
'        If Val(txtMonto.Text) = 0 Then
'            cmdGrabar.Enabled = False
'        Else
'            If gbITFAplica Then
'                Me.lblITF.Caption = Format(fgITFCalculaImpuesto(CCur(txtMonto.Text)), "#,##0.00")
'            End If
'
            vMonto = GetLimiteMonto(nmoneda)
           
'            If CDbl(Val(txtMonto.Text)) < vMonto Then
'               vNComision = GetComisionMayor()
'               vNComision = nComisionPorcentMaynas
'               If nMoneda = gMonedaNacional Then
'                    lblComision.Caption = Format(((Val(txtMonto.Text) * vNComision)), "#,##0.00")
'                    If Me.lblComision.Caption < nMontoMinimoComMaynas Then
'                        Me.lblComision.Caption = Format(nMontoMinimoComMaynas, "#,##0.00")
'                    End If
'               Else
'                    lblComision.Caption = Format(((txtMonto.value * vNComision)) * vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta), "#,##0.00")
'                    If Me.lblComision.Caption < (nMontoMinimoComMaynas / vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta)) Then
'                        Me.lblComision.Caption = Format((nMontoMinimoComMaynas / vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta)), "#,##0.00")
'                    End If
               
               'End If
'            Else
'                    'vNComision = GetValorComision()
'                    'lblComision.Caption = Format(vNComision, "#0.00")
'                    MsgBox "El Monto excede lo maximo permitido " & vMonto
'                    Set vTipoCambio = Nothing
'                    Exit Sub
'            End If
        
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
'        End If
        Set vTipoCambio = Nothing

        'txtSaldo.SetFocus
    End If
End Sub

Private Sub txtSaldo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGlosa.SetFocus
End If
End Sub
