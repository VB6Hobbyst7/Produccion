VERSION 5.00
Begin VB.Form frmColPOpeCMACLlam 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   Icon            =   "frmColPOpeCMACLlam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5580
      TabIndex        =   7
      Top             =   2610
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5580
      TabIndex        =   6
      Top             =   2160
      Width           =   990
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5580
      TabIndex        =   8
      Top             =   3060
      Width           =   990
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
      Height          =   2295
      Left            =   60
      TabIndex        =   10
      Top             =   1560
      Width           =   5490
      Begin VB.Frame Frame1 
         Caption         =   "Glosa"
         Height          =   1215
         Left            =   90
         TabIndex        =   17
         Top             =   990
         Width           =   2565
         Begin VB.TextBox txtGlosa 
            Height          =   870
            Left            =   60
            TabIndex        =   4
            Top             =   225
            Width           =   2415
         End
      End
      Begin VB.Frame fraDocumento 
         Caption         =   "Documento"
         Height          =   735
         Left            =   90
         TabIndex        =   15
         Top             =   225
         Width           =   2580
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
            Left            =   180
            TabIndex        =   3
            Top             =   270
            Width           =   2250
         End
      End
      Begin VB.Frame fraMonto 
         Height          =   1215
         Left            =   2685
         TabIndex        =   13
         Top             =   990
         Width           =   2715
         Begin SICMACT.EditMoney txtMonto 
            Height          =   345
            Left            =   1065
            TabIndex        =   5
            Top             =   720
            Width           =   1515
            _ExtentX        =   2672
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
            Height          =   330
            Left            =   1065
            TabIndex        =   20
            Top             =   315
            Width           =   1545
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Comisión S/:"
            Height          =   195
            Left            =   90
            TabIndex        =   19
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
            Height          =   195
            Left            =   90
            TabIndex        =   18
            Top             =   765
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
      TabIndex        =   9
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
         Top             =   660
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
         Top             =   1020
         Width           =   4215
      End
      Begin VB.Label Label6 
         Caption         =   "Moneda"
         Height          =   255
         Left            =   180
         TabIndex        =   16
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
         ForeColor       =   &H8000000D&
         Height          =   675
         Left            =   3600
         TabIndex        =   14
         Top             =   240
         Width           =   1875
      End
      Begin VB.Image imagen 
         Height          =   480
         Index           =   0
         Left            =   5640
         Picture         =   "frmColPOpeCMACLlam.frx":030E
         Top             =   420
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   1110
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
         Height          =   195
         Left            =   180
         TabIndex        =   11
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
Dim fnVarOpeCod As Long
Dim fsVarOpeDesc As String
Dim fsVarPersCodCMAC As String
Dim fsVarNombreCMAC As String
Dim fnVarMoneda As Moneda

Public Sub Inicio(ByVal pnOpeCod As Long, ByVal psOpeDesc As String, _
        ByVal psPersCodCMAC As String, ByVal psNomCmac As String, _
        Optional nComision As Double = 0)

fnVarOpeCod = pnOpeCod
fsVarOpeDesc = psOpeDesc
fsVarPersCodCMAC = psPersCodCMAC
fsVarNombreCMAC = psNomCmac
lblMensaje = fsVarNombreCMAC & Chr$(13) & fsVarOpeDesc

If pnOpeCod = 107001 Then
    Me.Caption = "Créditos CMAC Llamada - " & psOpeDesc
End If


Select Case fnVarOpeCod
    Case geColPRenDEOtCj
        txtDocumento.Visible = True
End Select
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
     vMonto = GetLimiteMonto(fnVarMoneda)
   
    If txtMonto.value > vMonto Then
        vNComision = GetComisionMayor()
       If fnVarMoneda = gMonedaNacional Then
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
fnVarMoneda = Right(cboMoneda.Text, 2)
If fnVarMoneda = gMonedaNacional Then
    txtMonto.BackColor = &HC0FFFF
Else
    txtMonto.BackColor = &HC0FFC0
End If

Dim vMonto As Double
Dim vNComision As Double
Dim vTipoCambio As COMDConstSistema.NCOMTipoCambio
Set vTipoCambio = New COMDConstSistema.NCOMTipoCambio

If txtMonto.value = 0 Then
    cmdGrabar.Enabled = False
Else
     vMonto = GetLimiteMonto(fnVarMoneda)
   
    If txtMonto.value > vMonto Then
        vNComision = GetComisionMayor()
       If fnVarMoneda = gMonedaNacional Then
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

Private Sub cmdGrabar_Click()
Dim lsCuenta As String, lsCliente As String
Dim lnMonto As Currency
Dim lsDocumento As String
Dim lnMoneda As Integer

Dim lsmensaje As String
Dim lsBoleta As String
Dim lsBoletaITF As String

Dim loLavDinero As frmMovLavDinero
Set loLavDinero = New frmMovLavDinero


lsCuenta = Trim(txtCuenta)
lsCliente = Trim(txtCliente)
lnMonto = txtMonto.value
lsDocumento = txtDocumento
If lsCuenta = "" Then
    MsgBox "Debe digitar un número de cuenta válido", vbInformation, "Aviso"
    txtCuenta.SetFocus
    Exit Sub
End If
If lsCliente = "" Then
    MsgBox "Debe digitar el nombre del cliente", vbInformation, "Aviso"
    txtCliente.SetFocus
    Exit Sub
End If
If lnMonto = 0 Then
    MsgBox "Debe colocar un monto mayor a cero", vbInformation, "Aviso"
    txtMonto.SetFocus
    Exit Sub
End If

'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarOpe As COMNColoCPig.NCOMColPContrato
Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lsIFTipo As String
Dim lsGlosa As String


lsIFTipo = Format$(gTpoIFCmac, "00")
lsGlosa = Trim(txtGlosa.Text)
lsDocumento = Trim(txtDocumento.Text)
lnMonto = txtMonto.value

If cboMoneda.ListIndex = 0 Then
    lnMoneda = 1
Else
    lnMoneda = 2
End If

If MsgBox(" Desea Grabar la Operación - " & fsVarOpeDesc & " ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios
    Dim nMontoLavDinero As Double, nTC As Double
    Dim sPersLavDinero As String

    Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
    If Not clsExo.EsCuentaExoneradaLavadoDinero(lsCuenta) Or Len(lsCuenta) < 18 Then
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Set clsExo = Nothing
        sPersLavDinero = ""
        nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
        Set clsLav = Nothing
        If lnMoneda = gMonedaNacional Then
            Dim clsTC As COMDConstSistema.NCOMTipoCambio
            Set clsTC = New COMDConstSistema.NCOMTipoCambio
            nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
            Set clsTC = Nothing
        Else
            nTC = 1
        End If
        If lnMonto >= Round(nMontoLavDinero * nTC, 2) Then
            'By Capi 1402208
            Call IniciaLavDinero(loLavDinero)
            'ALPA 20081009****************************************************
            'sperslavdinero = loLavDinero.Inicia(, , , , False, True, CDbl(txtMonto.Text), lsCuenta, Mid(Me.Caption, 15), False, "", , , , , lnMoneda)
            sPersLavDinero = loLavDinero.Inicia(, , , , False, True, CDbl(txtMonto.Text), lsCuenta, Mid(Me.Caption, 15), False, "", , , , , lnMoneda, , gnTipoREU, gnMontoAcumulado, gsOrigen)
            '*****************************************************************
            If loLavDinero.OrdPersLavDinero = "" Then Exit Sub
            'End

            'sperslavdinero = IniciaLavDinero()
            'If sperslavdinero = "" Then Exit Sub
        End If
    Else
        Set clsExo = Nothing
    End If
    
    
    Dim oCap As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim sCuentaAho As String
    Set oCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    sCuentaAho = oCap.GetCuentaAbonoIF(fsVarPersCodCMAC, fnVarMoneda)
    lsCuenta = sCuentaAho
    Set oCap = Nothing
    
    If sCuentaAho = "" Then
        MsgBox "No Existe cuenta de regularización.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    cmdGrabar.Enabled = False
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarOpe = New COMNColoCPig.NCOMColPContrato
            'Grabar la operacion
            Select Case fnVarOpeCod
                Case geColPRenDEOtCj
                    
                    Call loGrabarOpe.nOpeCMACLlamadaCredPignoraticio(fnVarOpeCod, lsFechaHoraGrab, _
                         lsMovNro, lsGlosa, fsVarPersCodCMAC, lsIFTipo, fnVarMoneda, lsCuenta, lsDocumento, lnMonto, False, CDbl(Val(lblComision.Caption)), gMonedaNacional, sCuentaAho, gsNomAge, sLpt, , lsmensaje, lsBoleta, lsBoletaITF, , gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
                         
                Case geColPCanDEOtCj
                
                    Call loGrabarOpe.nOpeCMACLlamadaCredPignoraticio(fnVarOpeCod, lsFechaHoraGrab, _
                         lsMovNro, lsGlosa, fsVarPersCodCMAC, lsIFTipo, fnVarMoneda, lsCuenta, lsDocumento, lnMonto, False, CDbl(Val(lblComision.Caption)), gMonedaNacional, sCuentaAho, gsNomAge, sLpt, , lsmensaje, lsBoleta, lsBoletaITF, , gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
                Case "107001"
                'Llamadas de credito
                    Call loGrabarOpe.nOpeCMACLlamadaCredPignoraticio(fnVarOpeCod, lsFechaHoraGrab, _
                         lsMovNro, lsGlosa, fsVarPersCodCMAC, lsIFTipo, fnVarMoneda, lsCuenta, lsDocumento, lnMonto, False, CDbl(Val(lblComision.Caption)), gMonedaNacional, sCuentaAho, gsNomAge, sLpt, lsCliente, sPersLavDinero, lsmensaje, lsBoleta, lsBoletaITF, gbImpTMU, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnMovNro)
            End Select
                'ALPA 20081010***********
            If gnMovNro > 0 Then
                 'Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen)
                 Call loLavDinero.InsertarLavDinero(loLavDinero.TitPersLavDinero, , , gnMovNro, loLavDinero.BenPersLavDinero, loLavDinero.TitPersLavDinero, loLavDinero.OrdPersLavDinero, loLavDinero.ReaPersLavDinero, loLavDinero.BenPersLavDinero, loLavDinero.VisPersLavDinero, gnTipoREU, gnMontoAcumulado, gsOrigen, loLavDinero.BenPersLavDinero2, loLavDinero.BenPersLavDinero3, loLavDinero.BenPersLavDinero4) 'JACA 20110224
            End If
            ' verificar AVMM
'            If Trim(lsboleta) <> "" Then
'                nFicSal = FreeFile
'                Open sLpt For Output As nFicSal
'                    Print #nFicSal, lsboleta
'                    Print #nFicSal, ""
'                Close #nFicSal
'            End If
'
'            If Trim(lsboletaitf) <> "" Then
'                nFicSal = FreeFile
'                Open sLpt For Output As nFicSal
'                    Print #nFicSal, lsboletaitf
'                    Print #nFicSal, ""
'                Close #nFicSal
'            End If

        Set loGrabarOpe = Nothing
        Set loLavDinero = Nothing
        

Dim nVez As Integer
nVez = 1
        'Impresión
        Call ImprimirRecibo(nVez, lsCuenta)
        
        Do While True
            If MsgBox("Desea reimprimir boletas de regularizacion?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                nVez = nVez + 1
                Call ImprimirRecibo(nVez, lsCuenta)
            Else
                Exit Do
            End If
        Loop
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

Private Sub IniciaLavDinero(poLavDinero As frmMovLavDinero)

Dim sPersCod As String
Dim sNombre As String
Dim sDireccion As String
Dim sDocId As String
Dim nMonto As Double
Dim sCuenta As String
Dim n_Moneda As Integer

If cboMoneda.ListIndex = 0 Then
    n_Moneda = 1
ElseIf cboMoneda.ListIndex = 1 Then
    n_Moneda = 2
End If

    poLavDinero.TitPersLavDinero = "" 'LblCodCli.Caption
    poLavDinero.TitPersLavDineroNom = "" 'LblNomCli.Caption
    poLavDinero.TitPersLavDineroDir = "" 'LblCliDirec.Caption
    poLavDinero.TitPersLavDineroDoc = "" 'LblDocNat.Caption
    
    'sPersCod = "" 'LblCodCli.Caption
    'sNombre = "" 'LblNomCli.Caption
    'sDireccion = "" 'LblCliDirec.Caption
    'sDocId = "" 'LblDocNat.Caption
    
    
    sCuenta = txtCuenta.Text

    nMonto = CDbl(txtMonto.Text)
    'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, True, nMonto, sCuenta, Trim(fnVarOpeCod), True, "COLOCACIONES", , , , , n_Moneda)
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub


Private Sub txtCliente_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras(KeyAscii)
If KeyAscii = 13 Then
    If Me.txtDocumento.Enabled And Me.txtDocumento.Visible Then
        txtDocumento.SetFocus
    Else
        txtMonto.SetFocus
    End If
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
Private Function GetLimiteMonto(ByVal nMoneda As Moneda) As Double

Dim nValor As Double
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion

 nValor = 0
 nValor = oCap.GetCapParametro(IIf(nMoneda = gMonedaNacional, 2078, 2079))

Set oCap = Nothing

GetLimiteMonto = nValor

End Function

Private Function GetComisionMayor() As Double

Dim rsPar As New ADODB.Recordset
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
Set rsPar = oCap.GetTarifaParametro(fnVarOpeCod, gMonedaNacional, 2077)
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
Set rsPar = oCap.GetTarifaParametro(fnVarOpeCod, gMonedaNacional, gCostoOperacionCMACLlam)
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
             vMonto = GetLimiteMonto(fnVarMoneda)
           
            If txtMonto.value > vMonto Then
                vNComision = GetComisionMayor()
               If fnVarMoneda = gMonedaNacional Then
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
        If cmdGrabar.Enabled Then cmdGrabar.SetFocus
    End If
End Sub

Private Sub txtSaldo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGlosa.SetFocus
End If
End Sub

Private Sub ImprimirRecibo(ByVal nVez As Integer, ByVal psCuenta As String)
Dim N As COMNCaptaGenerales.NCOMCaptaMovimiento
Set N = New COMNCaptaGenerales.NCOMCaptaMovimiento

Dim lsBoleta As String
Dim nFicSal As String

lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
lsBoleta = lsBoleta & N.ImprimeBoletaCMAC("CREDITO CMAC LLAMADA", fsVarOpeDesc, Trim(txtMonto.Text), Trim(txtCliente.Text), Trim(psCuenta), "", 0, fnVarMoneda, "", 0, 0, False, False, , , fsVarNombreCMAC, , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodAge, , , gbImpTMU) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea

lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea

lsBoleta = lsBoleta & N.ImprimeBoletaCMAC("CREDITO CMAC LLAMADA", "Comision CMAC Llam", Trim(lblComision.Caption), Trim(txtCliente.Text), Trim(psCuenta), "", 0, gMonedaNacional, "", 0, 0, False, False, , , fsVarNombreCMAC, , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodAge, , , gbImpTMU) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
lsBoleta = lsBoleta & N.ImprimeBoletaCMACRegula("CREDITO CMAC LLAMADA", fsVarOpeDesc, Trim(txtMonto.Text), Trim(txtCliente.Text), Trim(psCuenta), "", 0, fnVarMoneda, "", 0, 0, False, False, , , fsVarNombreCMAC, , gdFecSis, gsNomAge, gsCodUser, sLpt, , gsCodAge, False, txtCuenta.Text) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
If Trim(lsBoleta) <> "" Then
      nFicSal = FreeFile
      Open sLpt For Output As nFicSal
          Print #nFicSal, lsBoleta
          Print #nFicSal, ""
      Close #nFicSal
End If
Set N = Nothing
End Sub
