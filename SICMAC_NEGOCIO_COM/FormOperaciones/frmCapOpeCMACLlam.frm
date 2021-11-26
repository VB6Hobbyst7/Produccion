VERSION 5.00
Begin VB.Form frmCapOpeCMACLlam 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "frmCapOpeCMACLlam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   4605
      TabIndex        =   8
      Top             =   4350
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5685
      TabIndex        =   9
      Top             =   4350
      Width           =   1000
   End
   Begin VB.Frame fraMovimiento 
      Caption         =   "Movimiento"
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
      Height          =   2790
      Left            =   60
      TabIndex        =   11
      Top             =   1515
      Width           =   6675
      Begin VB.Frame fraGlosa 
         Caption         =   "Glosa"
         Height          =   1725
         Left            =   90
         TabIndex        =   22
         Top             =   960
         Width           =   2925
         Begin VB.TextBox txtGlosa 
            Height          =   1365
            Left            =   75
            TabIndex        =   7
            Top             =   240
            Width           =   2730
         End
      End
      Begin VB.Frame fraDocumento 
         Height          =   735
         Left            =   90
         TabIndex        =   18
         Top             =   225
         Width           =   6480
         Begin VB.TextBox txtDocumento 
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
            Left            =   3420
            TabIndex        =   4
            Top             =   240
            Width           =   1155
         End
         Begin VB.TextBox txtExtracto 
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
            MaxLength       =   4
            TabIndex        =   3
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblEtqDocumento 
            AutoSize        =   -1  'True
            Caption         =   "Orden Pago :"
            Height          =   195
            Left            =   2220
            TabIndex        =   20
            Top             =   330
            Width           =   945
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "N° Extracto :"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   330
            Width           =   900
         End
      End
      Begin VB.Frame fraMonto 
         Height          =   1725
         Left            =   3015
         TabIndex        =   14
         Top             =   960
         Width           =   3570
         Begin SICMACT.EditMoney txtSaldo 
            Height          =   345
            Left            =   1095
            TabIndex        =   6
            Top             =   840
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney txtMonto 
            Height          =   345
            Left            =   1095
            TabIndex        =   5
            Top             =   1245
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.Label lblITF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2730
            TabIndex        =   26
            Top             =   435
            Width           =   750
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "ITF"
            Height          =   180
            Left            =   2910
            TabIndex        =   25
            Top             =   150
            Width           =   480
         End
         Begin VB.Label lblComision 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1095
            TabIndex        =   24
            Top             =   405
            Width           =   1635
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Comisión S/:"
            Height          =   195
            Left            =   90
            TabIndex        =   23
            Top             =   495
            Width           =   900
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Saldo :"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   915
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   540
         End
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
      ForeColor       =   &H8000000D&
      Height          =   1455
      Left            =   60
      TabIndex        =   10
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
         Left            =   960
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
         TabIndex        =   21
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
         ForeColor       =   &H00FF0000&
         Height          =   675
         Left            =   3840
         TabIndex        =   17
         Top             =   240
         Width           =   1635
      End
      Begin VB.Image imagen 
         Height          =   480
         Index           =   0
         Left            =   5775
         Picture         =   "frmCapOpeCMACLlam.frx":030A
         Top             =   450
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   1110
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta :"
         Height          =   195
         Left            =   180
         TabIndex        =   12
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
'By capi 23022009
Dim lsMensajeValidacion As String
'
Dim nMontoMinimoComMaynas As Double
Dim nComisionPorcentMaynas As Double
Public Sub Inicia(ByVal nOpe As CaptacOperacion, ByVal sDescOpe As String, _
        ByVal sCodCMAC As String, ByVal sNomCMAC As String, Optional nComision As Double = 0)

sDescOperacion = sDescOpe
Me.Caption = "Captaciones - Operaciones CMACs LLamada - " & sDescOperacion
nOperacion = nOpe
sPersCodCMAC = sCodCMAC
sNombreCMAC = sNomCMAC
lblMensaje = sNombreCMAC & Chr$(13) & sDescOperacion
Select Case nOperacion
    Case gCMACOTAhoDepChq
        lblEtqDocumento.Visible = True
        txtDocumento.Visible = True
        lblEtqDocumento = "Cheque N° :"
    Case gCMACOTAhoRetOP
        lblEtqDocumento.Visible = True
        txtDocumento.Visible = True
        lblEtqDocumento = "Orden Pago N° :"
    Case Else
        lblEtqDocumento.Visible = False
        txtDocumento.Visible = False
End Select

If nOperacion = gCMACOTAhoDepChq Or nOperacion = gCMACOTAhoDepEfec Then
    txtMonto.ForeColor = &HC00000
ElseIf nOperacion = gCMACOTAhoRetEfec Or nOperacion = gCMACOTAhoRetOP Then
    txtMonto.ForeColor = &HC0&

End If

cboMoneda.AddItem "SOLES" & Space(100) & gMonedaNacional
cboMoneda.AddItem "DOLARES" & Space(100) & gMonedaExtranjera
cboMoneda.ListIndex = 0
lblComision = Format$(nComision, "#,##0.00")
Me.Show 1
End Sub

Private Sub cboMoneda_Change()
Dim vMonto As Double
Dim vNComision As Double
Dim vTipoCambio As COMDConstSistema.NCOMTipoCambio
Set vTipoCambio = New COMDConstSistema.NCOMTipoCambio

If txtMonto.value = 0 Then
    cmdGrabar.Enabled = False
Else
'    If gbITFAplica Then
'        Me.lblITF.Caption = Format(fgITFCalculaImpuesto(CCur(txtMonto.Text)), "#,##0.00")
'    End If
    
    vMonto = GetLimiteMonto(nmoneda)
   
    If CDbl(txtMonto.value) > vMonto Then
       vNComision = GetComisionMayor()
       If nmoneda = gMonedaNacional Then
            lblComision.Caption = Format(((txtMonto.value * vNComision) * 0.01), "#,##0.00")
       Else
            lblComision.Caption = Format(((txtMonto.value * vNComision) * 0.01) * vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta), "#,##0.00")
       End If
    Else
            vNComision = GetValorComision()
            lblComision.Caption = Format(vNComision, "#0.00")
    End If

    cmdGrabar.Enabled = True
End If
Set vTipoCambio = Nothing
End Sub

Private Sub cboMoneda_Click()
'*********modificado para que solo cobre comision en SOLES


nmoneda = Right(cboMoneda.Text, 2)
If nmoneda = gMonedaNacional Then
    txtMonto.BackColor = &HC0FFFF
    txtSaldo.BackColor = &HC0FFFF
Else
    txtMonto.BackColor = &HC0FFC0
    txtSaldo.BackColor = &HC0FFC0
End If
Dim vMonto As Double
Dim vNComision As Double
Dim vTipoCambio As COMDConstSistema.NCOMTipoCambio
Set vTipoCambio = New COMDConstSistema.NCOMTipoCambio

If txtMonto.value = 0 Then
    cmdGrabar.Enabled = False
Else
    vMonto = GetLimiteMonto(nmoneda)
    If CDbl(txtMonto.value) > vMonto Then
       vNComision = GetComisionMayor()
       If nmoneda = gMonedaNacional Then
            lblComision.Caption = Format(((txtMonto.value * vNComision) * 0.01), "#,##0.00")
       Else
            lblComision.Caption = Format(((txtMonto.value * vNComision) * 0.01) * vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta), "#,##0.00")
       End If
    Else
            vNComision = GetValorComision()
            lblComision.Caption = Format(vNComision, "#0.00")
    End If

    cmdGrabar.Enabled = True
End If
Set vTipoCambio = Nothing
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCuenta.SetFocus
End If
End Sub

Private Sub CmdGrabar_Click()
Dim sCuenta As String, sCliente As String
Dim nMonto As Double, nSaldo As Double
Dim nExtracto As Long
Dim sDocumento As String
Dim loLavDinero As frmMovLavDinero
Set loLavDinero = New frmMovLavDinero

'By capi 23022009
lsMensajeValidacion = ""
'

nExtracto = Val(txtExtracto.Text)
sCuenta = Trim(txtCuenta)

sCliente = Trim(txtCliente)
nMonto = txtMonto.value
nSaldo = txtSaldo.value
sDocumento = txtDocumento

If sCuenta = "" Then
    MsgBox "Debe digitar un número de cuenta válido", vbInformation, "Aviso"
    txtCuenta.SetFocus
    Exit Sub
End If
If sCliente = "" Then
    MsgBox "Debe digitar el nombre del cliente", vbInformation, "Aviso"
    txtCliente.SetFocus
    Exit Sub
End If
If nMonto = 0 Then
    MsgBox "Debe colocar un monto mayor a cero", vbInformation, "Aviso"
    txtMonto.SetFocus
    Exit Sub
End If
If nOperacion = gCMACOTAhoDepChq Or nOperacion = gCMACOTAhoRetOP Then
    If sDocumento = "" Then
        MsgBox "Debe digitar un documento válido", vbInformation, "Aviso"
        txtDocumento.SetFocus
        Exit Sub
    End If
End If

If MsgBox("Desea Grabar la Operacion??", vbQuestion + vbYesNo, "Aviso") = vbYes Then

    If nOperacion = gCMACOTAhoDepEfec Or nOperacion = gCMACOTAhoRetEfec Then

        Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios
        Dim nMontoLavDinero As Double, nTC As Double
        Dim sPersLavDinero As String, sReaPersLavDinero As String, sBenPersLavDinero As String

        Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
        If Not clsExo.EsCuentaExoneradaLavadoDinero(sCuenta) Or Len(sCuenta) < 18 Then
            Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
            Set clsExo = Nothing
            sPersLavDinero = ""
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            If nmoneda = gMonedaNacional Then
                Dim clsTC As COMDConstSistema.NCOMTipoCambio
                Set clsTC = New COMDConstSistema.NCOMTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
            Else
                nTC = 1
            End If
            If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                'By Capi 1402208
                    Call IniciaLavDinero(loLavDinero)
                    'ALPA 20081009************************************************************
                    sPersLavDinero = loLavDinero.Inicia(, , , , False, True, nMonto, sCuenta, Mid(Me.Caption, 15), True, "", , , , , nmoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
                    If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
                'End
            End If
        Else
            Set clsExo = Nothing
        End If

    End If
    

    Dim sGlosa As String, sMovNro As String, sCtaAbono As String
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
    
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    sCtaAbono = oCap.GetCuentaAbonoIF(sPersCodCMAC, nmoneda)
    sCuenta = sCtaAbono
    Set oCap = Nothing
    
    Dim lsBoleta As String
    Dim nFicSal As Integer
    
    If Len(Trim(sCtaAbono)) = 0 Then
        MsgBox "No existe cuenta con la cual hacer la regularización", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Select Case nOperacion
        Case gCMACOTAhoDepEfec
            'Entra
            clsCap.CapOpeAhoCMACLlamada sMovNro, nmoneda, sGlosa, gCMACOTAhoDepEfec, nExtracto, sDescOperacion, nMonto, sCuenta, nSaldo, sPersCodCMAC, sNombreCMAC, sCliente, , , sCtaAbono, gsNomAge, sLpt, CDbl(Val(lblComision.Caption)), gMonedaNacional, sPersLavDinero, CCur(Val(Me.lblITF.Caption)), sBenPersLavDinero, lsBoleta, txtCuenta.Text, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro
        Case gCMACOTAhoDepChq
            clsCap.CapOpeAhoCMACLlamada sMovNro, nmoneda, sGlosa, gCMACOTAhoDepChq, nExtracto, sDescOperacion, nMonto, sCuenta, nSaldo, sPersCodCMAC, sNombreCMAC, sCliente, TpoDocCheque, sDocumento, , gsNomAge, sLpt, CDbl(Val(lblComision.Caption)), gMonedaNacional, CCur(Val(Me.lblITF.Caption)), , sBenPersLavDinero, lsBoleta, txtCuenta.Text, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro
        Case gCMACOTAhoRetEfec
            'By capi 23022009
            'clsCap.CapOpeAhoCMACLlamada sMovNro, nMoneda, sGlosa, gCMACOTAhoRetEfec, nExtracto, sDescOperacion, nMonto, sCuenta, nSaldo, sPersCodCMAC, sNombreCMAC, sCliente, , , sCtaAbono, gsNomAge, sLpt, CDbl(Val(lblComision.Caption)), gMonedaNacional, sPersLavDinero, CCur(Val(Me.lblITF.Caption)), sBenPersLavDinero, lsBoleta, txtCuenta.Text, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro
            clsCap.CapOpeAhoCMACLlamada sMovNro, nmoneda, sGlosa, gCMACOTAhoRetEfec, nExtracto, sDescOperacion, nMonto, sCuenta, nSaldo, sPersCodCMAC, sNombreCMAC, sCliente, , , sCtaAbono, gsNomAge, sLpt, CDbl(Val(lblComision.Caption)), gMonedaNacional, sPersLavDinero, CCur(Val(Me.lblITF.Caption)), sBenPersLavDinero, lsBoleta, txtCuenta.Text, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro, lsMensajeValidacion
            '
        Case gCMACOTAhoRetOP
            clsCap.CapOpeAhoCMACLlamada sMovNro, nmoneda, sGlosa, gCMACOTAhoRetOP, nExtracto, sDescOperacion, nMonto, sCuenta, nSaldo, sPersCodCMAC, sNombreCMAC, sCliente, TpoDocOrdenPago, sDocumento, sCtaAbono, gsNomAge, sLpt, CDbl(Val(lblComision.Caption)), gMonedaNacional, CCur(Val(Me.lblITF.Caption)), , sBenPersLavDinero, lsBoleta, txtCuenta.Text, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro
    End Select
    'By capi 23022009 para visualizar mensaje de validacion
    If lsMensajeValidacion <> "" Then
        MsgBox lsMensajeValidacion, vbInformation, "Aviso"
        Exit Sub
    End If
    
    'End by
    
    'ALPA 20081010
    If gnMovNro > 0 Then
     'Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, sBenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
      Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, sBenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
    End If
    Do
        If Trim(lsBoleta) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsBoleta
                Print #nFicSal, ""
            Close #nFicSal
        End If
    Loop Until MsgBox("¿Desea Re-Imprimir Boletas ?", vbQuestion + vbYesNo, "Aviso") = vbNo
    
    gVarPublicas.LimpiaVarLavDinero
    Set loLavDinero = Nothing
    Set clsCap = Nothing
    Unload Me
End If
End Sub


Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero)

Dim sPersCod As String
Dim sNombre As String
Dim sDireccion As String
Dim sDocId As String
Dim nMonto As Double
Dim sCuenta As String
Dim n_Moneda As Integer

sPersCod = "" 'LblCodCli.Caption
sNombre = "" 'LblNomCli.Caption
sDireccion = "" 'LblCliDirec.Caption
sDocId = "" 'LblDocNat.Caption

poLavDinero.TitPersLavDinero = ""
poLavDinero.TitPersLavDineroNom = ""
poLavDinero.TitPersLavDineroDir = ""
poLavDinero.TitPersLavDineroDoc = ""


If cboMoneda.ListIndex = 0 Then
    n_Moneda = 1
ElseIf cboMoneda.ListIndex = 1 Then
    n_Moneda = 2
End If
    
    sCuenta = txtCuenta.Text

nMonto = CDbl(txtMonto.Text)
'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, True, nMonto, sCuenta, Trim(nOperacion), True, "COLOCACIONES", , , , , n_Moneda)
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
nMontoMinimoComMaynas = GetMontoMinimoMaynas
nComisionPorcentMaynas = GetPorcentMaynas
End Sub

Private Function GetMontoMinimoMaynas() As Double

Dim nValor As Double
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion

 nValor = 0
 nValor = oCap.GetCapParametro(2048)

Set oCap = Nothing
GetMontoMinimoMaynas = nValor
End Function
Private Function GetPorcentMaynas() As Double

Dim nValor As Double
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion

 nValor = 0
 nValor = oCap.GetCapParametro(2049)

Set oCap = Nothing
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
    txtExtracto.SetFocus
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

Private Sub txtExtracto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtDocumento.Visible Then
        txtDocumento.SetFocus
    Else
        txtMonto.SetFocus
    End If
    Exit Sub
End If
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    If cmdGrabar.Enabled = True Then
        cmdGrabar.SetFocus
    End If
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
''    If gbITFAplica Then
''        Me.lblITF.Caption = Format(fgITFCalculaImpuesto(CCur(txtMonto.Text)), "#,##0.00")
''    End If
'
'    vMonto = GetLimiteMonto(nMoneda)
'
'    If CDbl(txtMonto.value) > vMonto Then
'       vNComision = GetComisionMayor()
'       If nMoneda = gMonedaNacional Then
'            lblComision.Caption = Format(((txtMonto.value * vNComision) * 0.01), "#,##0.00")
'       Else
'            lblComision.Caption = Format(((txtMonto.value * vNComision) * 0.01) * vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta), "#,##0.00")
'       End If
'    Else
'            vNComision = GetValorComision()
'            lblComision.Caption = Format(vNComision, "#0.00")
'    End If
'
'    cmdGrabar.Enabled = True
'End If
'Set vTipoCambio = Nothing
End Sub
Private Function GetLimiteMonto(ByVal nmoneda As Moneda) As Double

Dim nValor As Double
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion

 nValor = 0
 nValor = oCap.GetCapParametro(IIf(nmoneda = gMonedaNacional, 2078, 2079))

Set oCap = Nothing
GetLimiteMonto = nValor
End Function

Private Function GetComisionMayor() As Double

Dim rsPar As New ADODB.Recordset
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set rsPar = oCap.GetTarifaParametro(nOperacion, gMonedaNacional, 2077)
Set oCap = Nothing
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
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
Set rsPar = oCap.GetTarifaParametro(nOperacion, gMonedaNacional, gCostoOperacionCMACLlam)
Set oCap = Nothing
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
   Dim vTipoCambio As COMDConstSistema.NCOMTipoCambio
 
   If KeyAscii = 13 Then
        Set vTipoCambio = New COMDConstSistema.NCOMTipoCambio
        
        If txtMonto.value = 0 Then
            cmdGrabar.Enabled = False
        Else
        '    If gbITFAplica Then
        '        Me.lblITF.Caption = Format(fgITFCalculaImpuesto(CCur(txtMonto.Text)), "#,##0.00")
        '    End If
            
            vMonto = GetLimiteMonto(nmoneda)
           
            If CDbl(txtMonto.value) < vMonto Then
               'vNComision = GetComisionMayor()
               vNComision = nComisionPorcentMaynas
               'If nMoneda = gMonedaNacional Then
                    lblComision.Caption = Format(((txtMonto.value * vNComision)), "#,##0.00")
                    If Me.lblComision.Caption < nMontoMinimoComMaynas Then
                        Me.lblComision.Caption = Format(nMontoMinimoComMaynas, "#,##0.00")
                    End If
'               Else
'                    lblComision.Caption = Format(((txtMonto.value * vNComision)) * vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta), "#,##0.00")
'                    If Me.lblComision.Caption < (nMontoMinimoComMaynas / vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta)) Then
'                        Me.lblComision.Caption = Format((nMontoMinimoComMaynas / vTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta)), "#,##0.00")
'                    End If
               
               'End If
            Else
                    'vNComision = GetValorComision()
                    'lblComision.Caption = Format(vNComision, "#0.00")
                    MsgBox "El Monto excede lo maximo permitido " & vMonto
                    Set vTipoCambio = Nothing
                    Exit Sub
            End If
        
            cmdGrabar.Enabled = True
        End If
        Set vTipoCambio = Nothing

        txtSaldo.SetFocus
    End If
End Sub

Private Sub txtSaldo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGlosa.SetFocus
End If
End Sub
