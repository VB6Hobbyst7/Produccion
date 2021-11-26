VERSION 5.00
Begin VB.Form frmDEConversionReConversion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dinero Electrónico"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDEConversionReConversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleMode       =   0  'User
   ScaleWidth      =   5832.972
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmTipoPago 
      Caption         =   "Datos de ahorro y tipo de pago"
      Height          =   975
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   6675
      Begin VB.ComboBox CmbForPag 
         Height          =   345
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   360
         Width           =   1365
      End
      Begin SICMACT.ActXCodCta txtCuentaCargo 
         Height          =   375
         Left            =   2760
         TabIndex        =   22
         Top             =   340
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCta      =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblTipoPago 
         Caption         =   "Tipo Pago:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   450
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4200
      TabIndex        =   7
      Top             =   4680
      Width           =   1185
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   5520
      TabIndex        =   8
      Top             =   4680
      Width           =   1155
   End
   Begin VB.Frame fraTipoCambio 
      Height          =   1065
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   6675
      Begin VB.TextBox TxtMontoPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin SICMACT.EditMoney txtImporte 
         Height          =   360
         Left            =   1800
         TabIndex        =   5
         Top             =   160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label lblMonto 
         Caption         =   "Monto a Obtener:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label lblTipoCambio 
         Caption         =   "Monto a Convertir:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   255
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "S/"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "S/ - DE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame fraPersona 
      Caption         =   "Datos de la Persona"
      Height          =   2025
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6675
      Begin VB.TextBox txtMSISDN 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   2760
         MaxLength       =   9
         TabIndex        =   0
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtDNI 
         Height          =   315
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   2
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   2760
         TabIndex        =   3
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtDomicilio 
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "DNI"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre completo"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Domicilio (opcional)"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "N° de Móvil"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   -1680
      TabIndex        =   23
      Top             =   4680
      Width           =   3960
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   6285
   End
End
Attribute VB_Name = "frmDEConversionReConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnTipoMov As Integer ' 1 conversion - 2 reconversion
Dim lcPrdConcepto As Integer
Dim lsConceptoDesc As String
Dim lcTituloImp As String
'CTI7 OPEv2****************************************************************
Dim sNumTarj As String
Dim pnMoneda As Integer
Dim psPersCodTitularAhorroCargoDeposito As String
Dim psDNITitularAhorroCargoDeposito As String
Dim pnITF As Double
Dim pbEsMismoTitular As Boolean
Dim pnMontoPagarCargoDeposito As Double
Dim lbITFCtaExonerada As Boolean
Dim nmoneda As Integer
Dim pbOrdPag As Boolean
Dim lblTipoCuenta As String
Dim lblFirmas As String
Dim bValidaCantDep As Boolean
Dim nParCantDepLib As Integer
Dim nParMontoMinDepSol As Double
Dim nParMontoMinDepDol As Double
Dim nRespuesta As Integer 'CTI4 ERS0112020
Dim loVistoElectronico As frmVistoElectronico
Dim nRedondeoITF As Double ' BRGO 20110914
Public nProducto As COMDConstantes.Producto
'**************************************************************************
Private Sub cmdGuardar_Click()
    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII
    
    Dim oCajero As COMNCajaGeneral.NCOMCajero
    Dim ObjTcP As COMDConstSistema.DCOMTCEspPermiso
    Dim cNumeroMovil As String
    Dim cDNI As String
    Dim cPersNombre As String
    Dim cPersDireccion As String
    Dim cperscod As String
    
    Dim lsBoletaCargo  As String 'CTI7 OPEv2
    Dim MatDatosAho(14) As String 'CTI7 OPEv2
    Dim lsNombreClienteCargoDepositoCta As String 'CTI7 OPEv2
    Dim lsOpeDesc As String 'CTI7 OPEv2
    
    cNumeroMovil = Trim(txtMSISDN.Text)
    cDNI = Trim(txtDNI.Text)
    cPersNombre = UCase(Trim(txtNombre.Text))
    cPersDireccion = UCase(Trim(txtDomicilio.Text))
    
    Dim objPersona As COMDPersona.DCOMPersonas
    Set objPersona = New COMDPersona.DCOMPersonas
    Set oCajero = New COMNCajaGeneral.NCOMCajero
    
    If ValidaInterfaz = False Then Exit Sub
    
        Dim rsPersVerifica As Recordset
        Dim i As Integer
        Set rsPersVerifica = New Recordset
        Set rsPersVerifica = objPersona.ObtenerDatosPersonaXDNI(cDNI)
        If Not (rsPersVerifica.BOF And rsPersVerifica.EOF) Then
            cperscod = rsPersVerifica!cperscod
        End If
        
        'CTI7 OPEv2******************************************************************
        Dim clsCapN As New COMNCaptaGenerales.NCOMCaptaMovimiento
        lsOpeDesc = gsOpeDesc
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Or CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoDeposito Then
           
           If gsCodPersUser = psPersCodTitularAhorroCargoDeposito Then
                 MsgBox "Usted no puede hacer procesos de conversión/reconversión con su propia información", vbInformation, "¡Aviso!"
                 Exit Sub
           End If
             
             
           If Not clsCapN.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, pnMontoPagarCargoDeposito) Then
               If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
                    MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
               End If
               Exit Sub
           End If
           
           Dim nEstadoCaptacion As CaptacEstado
           
           nEstadoCaptacion = clsCapN.ObtieneEstadoCuenta(txtCuentaCargo.NroCuenta)
           
           If Not (nEstadoCaptacion = gCapEstActiva) And CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
                MsgBox "Cuenta a debitar NO está ACTIVA", vbInformation, "¡Aviso!"
                Exit Sub
           End If
           
            If Not (nEstadoCaptacion = gCapEstActiva Or nEstadoCaptacion = gCapEstBloqRetiro) And CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoDeposito Then
                MsgBox "Cuenta a depositar NO está ACTIVA", vbInformation, "¡Aviso!"
                Exit Sub
           End If
           
           Select Case CInt(Trim(Right(CmbForPag.Text, 10)))
               Case gColocTipoPagoCargoCta
                   lblTitulo = "CCONVERSION DINERO FISICO A DINERO ELECTRONICO"
                   lsOpeDesc = gsOpeDesc & "-CARGO A CUENTA"
               Case gColocTipoPagoDeposito
                   lblTitulo = "CONVERSION DINERO ELECTRONICO A DINERO FISICO"
                   lsOpeDesc = gsOpeDesc & "-DEPOSITO"
           End Select
        End If
        If Not ValidaFormaPago Then Exit Sub 'CTI7 OPEv2
        
        '*****************************************************************************
    If gsCodPersUser = cperscod Then
        MsgBox "Usted no puede hacer procesos de conversión/reconversión con su propia información", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    

    If MsgBox("¿Desea grabar la operación de conversión/reconversión?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        Dim nMonto As Double
        Dim lsBoleta As String
        Dim lsBoletaAhorro As String
        Dim nFicSal As Integer
        Dim cMovNro As String
        nMonto = txtImporte.value
        
        'If oCajero.GrabaConversionReconversionDE(cNumeroMovil, cDNI, cPersNombre, cPersDireccion, cperscod, nMonto, lcPrdConcepto, lsConceptoDesc, gsOpeCod, lnTipoMov, gdFecSis, gsCodAge, gsCodUser, cMovNro) = 0 Then
        If oCajero.GrabaConversionReconversionDE(cNumeroMovil, cDNI, cPersNombre, cPersDireccion, cperscod, nMonto, lcPrdConcepto, lsConceptoDesc, gsOpeCod, lnTipoMov, gdFecSis, gsCodAge, gsCodUser, cMovNro, CInt(Trim(Right(IIf(Trim(CmbForPag.Text) = "", "0000001", CmbForPag.Text), 10))), txtCuentaCargo.NroCuenta, MatDatosAho, gITF.gbITFAplica, pnITF, pnMontoPagarCargoDeposito, gsNomAge, lsBoletaAhorro, gbImpTMU, False) = 0 Then
            
            Dim oImp As COMNContabilidad.NCOMContImprimir
            Dim lsTexto As String
            Dim lbReimp As Boolean
            Set oImp = New COMNContabilidad.NCOMContImprimir
            
            lsBoleta = oImp.ImprimeBoletaConversionREConversion(lcTituloImp, "", cPersNombre, cPersDireccion, cDNI, _
            gsOpeCod, CCur(txtImporte), CCur(TxtMontoPagar), gsNomAge, cMovNro, sLpt, gsCodCMAC, gsNomCmac, gbImpTMU)
            lbReimp = True
            Do While lbReimp
                 If Trim(lsBoleta) <> "" Then
                    nFicSal = FreeFile
                    Open sLpt For Output As nFicSal
                        Print #nFicSal, lsBoleta
                        Print #nFicSal, ""
                    Close #nFicSal
                 End If
               
                If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                    lbReimp = False
                End If
            Loop
            
            If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Or CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoDeposito Then
                If Trim(lsBoletaAhorro) <> "" Then
                    lbReimp = True
                    Do While lbReimp
                         If Trim(lsBoletaAhorro) <> "" Then
                            nFicSal = FreeFile
                            Open sLpt For Output As nFicSal
                                Print #nFicSal, lsBoletaAhorro
                                Print #nFicSal, ""
                            Close #nFicSal
                         End If
                       
                        If MsgBox("Desea Reimprimir boleta de Operación de ahorro", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                            lbReimp = False
                        End If
                    Loop
                End If
            End If
            Set oImp = Nothing
            
            MsgBox "Operación realizada con éxito.", vbInformation, "Aviso"
            txtMSISDN.Text = ""
            txtDNI.Text = ""
            txtNombre.Text = ""
            txtDomicilio.Text = ""
            txtImporte = 0
            TxtMontoPagar = "0.00"
            txtImporte.Enabled = True
            Call IniciarControlesFormaPago
            'INICIO JHCU ENCUESTA 16-10-2019
            Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
            'FIN
        End If
        
        Set oCajero = Nothing
    End If
End Sub

Function ValidaInterfaz() As Boolean
   ValidaInterfaz = True
   If (txtMSISDN.Text = "") Then
        MsgBox "Número móvil no ingresado", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtMSISDN.SetFocus
        Exit Function
   ElseIf Len(txtMSISDN.Text) > 9 Or Len(txtMSISDN.Text) < 9 Then
        MsgBox "Número móvil no válido", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtMSISDN.SetFocus
        Exit Function
   End If
   
   If (txtDNI.Text = "") Then
        MsgBox "DNI no válido", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtDNI.SetFocus
        Exit Function
   ElseIf Len(txtDNI.Text) > 8 Or Len(txtDNI.Text) < 8 Then
        MsgBox "DNI no válido", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtDNI.SetFocus
        Exit Function
   End If
   
   If (txtNombre.Text = "") Then
        MsgBox "Nombre de la persona no ingresada", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtNombre.SetFocus
        Exit Function
   End If
   
   If Val(txtImporte) = 0 Then
        MsgBox "Importe de Operación no Ingresado", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtImporte.SetFocus
        Exit Function
    End If
    If Val(txtImporte) > 999 Then
        MsgBox "", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtImporte.SetFocus
        Exit Function
    End If
    If Val(TxtMontoPagar) = 0 Then
        MsgBox "Monto a Pagar no válido para Operación", vbInformation, "Aviso"
        ValidaInterfaz = False
        Exit Function
    End If

End Function
Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = gsOpeDesc
    txtImporte.psSoles False
    
    Select Case gsOpeCod
        Case COMDConstSistema.gOpeDEConversion
            lblTitulo.Caption = "CONVERSION DINERO FISICO A DINERO ELECTRONICO"
            lcTituloImp = "Conver. Dinero Fisico a Dinero Electr."
            Me.lblMonto = "Monto a Meter"
            lcPrdConcepto = COMDConstSistema.gConceptoConversion
            lnTipoMov = 1
            Call CargaControles(1)
        Case COMDConstSistema.gOpeDEReConversion
            lblTitulo = "CONVERSION DINERO ELECTRONICO A DINERO FISICO"
            lcTituloImp = "Conver. Dinero Electr. a Dinero Fisico"
            Me.lblMonto = "Monto a Sacar"
            lcPrdConcepto = COMDConstSistema.gConceptoReConversion
            lnTipoMov = 2
            Call CargaControles(2)
    End Select
    TxtMontoPagar = "0.00"
    lsConceptoDesc = lblTitulo.Caption
    Call IniciarControlesFormaPago
    CmbForPag.ListIndex = 0

End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNombre.SetFocus
        If Len(txtDNI) = 8 Then
            CmbForPag.Enabled = True
        Else
            CmbForPag.Enabled = True
        End If
    Else
        ValidarTeclaSoloNumeros KeyAscii
    End If
End Sub


Private Sub txtDNI_LostFocus()
    txtNombre.SetFocus
    If Len(txtDNI) = 8 Then
        CmbForPag.Enabled = True
    Else
        CmbForPag.Enabled = True
    End If
End Sub

Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtImporte.SetFocus
    End If
End Sub

Private Sub txtImporte_GotFocus()
    With txtImporte
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    Dim rs As ADODB.Recordset
    Dim ObjTc As COMDConstSistema.DCOMTipoCambioEsp

    If KeyAscii = 13 Then
        If txtImporte.Text = "" Or txtImporte.Text = 0 Then
            MsgBox "Ingrese el monto a Convertir", vbInformation, "AVISO"
            txtImporte.SetFocus
            Exit Sub
        End If
           
        TxtMontoPagar = Format(Val(txtImporte.value), "#,#0.00")
        Me.CmdGuardar.Enabled = True
        CmdGuardar.SetFocus
    End If
End Sub


Private Sub txtMSISDN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDNI.SetFocus
    Else
        If Not ValidarTeclaSoloNumeros(KeyAscii) Then
            Exit Sub
        End If
    End If
End Sub

Public Function ValidarTeclaSoloNumeros(KeyAscii As Integer) As Boolean
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 127 Or KeyAscii = 8) Then
        ValidarTeclaSoloNumeros = True
    Else
        KeyAscii = 0
        ValidarTeclaSoloNumeros = False
    End If
End Function
Public Function ValidarTeclaSoloLetras(KeyAscii As Integer) As Boolean
    If ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 127 Or KeyAscii = 8) Or KeyAscii = 32 Then
        ValidarTeclaSoloLetras = True
    Else
        KeyAscii = 0
        ValidarTeclaSoloLetras = False
    End If
End Function

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDomicilio.SetFocus
    Else
        If Not ValidarTeclaSoloLetras(KeyAscii) Then
            Exit Sub
        End If
    End If
End Sub

'CTI7 OPEv2*************************************************************************
'Private Sub IniciaCombo(ByRef cboConst As ComboBox, nCapConst As ConstanteCabecera, ByVal nTpoConversion As Integer)
'    Dim clsGen As COMDConstSistema.DCOMGeneral
'    Dim rsConst As New ADODB.Recordset
'    Set clsGen = New COMDConstSistema.DCOMGeneral
'    Set rsConst = clsGen.GetConstante(nCapConst)
'    Set clsGen = Nothing
'    Do While Not rsConst.EOF
'        cboConst.AddItem rsConst("cDescripcion") & space(100) & rsConst("nConsValor")
'        rsConst.MoveNext
'    Loop
'    If nTpoConversion = 1 Then
'     cboConst.AddItem "CARGO A CUENTA" & space(100) & 4
'    Else
'     cboConst.AddItem "DEPOSITO" & space(100) & 8
'    End If
'    cboConst.ListIndex = 0
'End Sub
Private Sub CargaControles(ByVal nTpoConversion As Integer)
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R = oCons.RecuperaConstantes(gColocTipoPago, , , 5)
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, CmbForPag)
    If nTpoConversion = 1 Then
     CmbForPag.AddItem "CARGO A CUENTA" & space(100) & 4
    Else
     CmbForPag.AddItem "DEPOSITO" & space(100) & 8
    End If
    Set loVistoElectronico = New frmVistoElectronico
    Exit Sub
ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub EstadoFormaPago(ByVal nFormaPago As Integer)

    txtCuentaCargo.NroCuenta = ""
    Select Case nFormaPago
        Case -1
            txtCuentaCargo.Visible = False
            CmdGuardar.Enabled = True
        Case gColocTipoPagoEfectivo
            txtCuentaCargo.Visible = False
            CmdGuardar.Enabled = True
        Case gColocTipoPagoCargoCta
            txtCuentaCargo.Visible = True
            txtCuentaCargo.Enabled = True
            txtCuentaCargo.CMAC = gsCodCMAC
            txtCuentaCargo.Prod = Trim(Str(gCapAhorros))
            CmdGuardar.Enabled = False
        Case gColocTipoPagoDeposito
            txtCuentaCargo.Visible = True
            txtCuentaCargo.Enabled = True
            txtCuentaCargo.CMAC = gsCodCMAC
            txtCuentaCargo.Prod = Trim(Str(gCapAhorros))
            CmdGuardar.Enabled = False
    End Select
End Sub
Private Sub CmbForPag_Click()
    EstadoFormaPago IIf(CmbForPag.ListIndex = -1, -1, CInt(Trim(Right(IIf(CmbForPag.Text = "", "-1", CmbForPag.Text), 10))))
    Dim sCuenta As String
    If CmbForPag.ListIndex <> -1 Then
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gCVTipoPagoEfectivo Then
     
            Dim oformVou As New frmCapRegVouDepBus
            Dim lnTipMot As Integer
            Dim sGlosa As String
            Dim sIF As String
            Dim sVaucher As String
            Dim sPersCod As String
            Dim sNombre As String
            Dim sDireccion As String
            Dim sDocumento As String
            lnTipMot = 15
            CmdGuardar.Enabled = True
            Call IniciarControlesCuentaCargoAbono
            txtCuentaCargo.Visible = False
        ElseIf CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            sCuenta = ""
            sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(gAhoCargoDineroElectronico), sNumTarj, 232, False)
            If Val(Mid(sCuenta, 6, 3)) <> "232" And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            txtCuentaCargo.Age = Mid(sCuenta, 4, 2)
            txtCuentaCargo.Prod = Mid(sCuenta, 6, 3)
            txtCuentaCargo.Cuenta = Mid(sCuenta, 9, 18)
            Call txtCuentaCargo_KeyPress(13)
            Call IniciarControlesCuentaCargoAbono
            txtCuentaCargo.Visible = True
            txtCuentaCargo.Enabled = True
            If Len(txtCuentaCargo.NroCuenta) <> 18 Then
                txtCuentaCargo.Age = ""
                txtCuentaCargo.Prod = gCapAhorros
                txtCuentaCargo.Cuenta = ""
                txtCuentaCargo.CMAC = gsCodCMAC
            End If
        ElseIf CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoDeposito Then
            sCuenta = ""
            sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(gAhoDepositoDineroElectronico), sNumTarj, 232, False)
            If Val(Mid(sCuenta, 6, 3)) <> "232" And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            txtCuentaCargo.Age = Mid(sCuenta, 4, 2)
            txtCuentaCargo.Prod = Mid(sCuenta, 6, 3)
            txtCuentaCargo.Cuenta = Mid(sCuenta, 9, 18)
            Call txtCuentaCargo_KeyPress(13)
            Call IniciarControlesCuentaCargoAbono
            txtCuentaCargo.Visible = True
            txtCuentaCargo.Enabled = True
            If Len(txtCuentaCargo.NroCuenta) <> 18 Then
                txtCuentaCargo.Age = ""
                txtCuentaCargo.Prod = gCapAhorros
                txtCuentaCargo.Cuenta = ""
                txtCuentaCargo.CMAC = gsCodCMAC
            End If
        End If
    End If
End Sub
Private Function ValidaFormaPago() As Boolean
Dim clsCap As New COMNCaptaGenerales.NCOMCaptaMovimiento
    ValidaFormaPago = False
    If CmbForPag.ListIndex = -1 Then
        MsgBox "No se ha seleccionado la forma de pago. Verifique.", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
    
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta And Len(txtCuentaCargo.NroCuenta) <> 18 Then
        MsgBox "Favor de ingresar la cuenta de ahorros a debitar.", vbInformation, "¡Aviso!"
        EnfocaControl CmbForPag
        Exit Function
    End If
        
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
        If Not clsCap.ValidaSaldoCuenta(txtCuentaCargo.NroCuenta, CDbl(TxtMontoPagar.Text)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA.", vbInformation, "¡Aviso!"
            Exit Function
        End If
    End If
    
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoDeposito And Len(txtCuentaCargo.NroCuenta) <> 18 Then
        MsgBox "Favor de ingresar la cuenta de ahorros a depositar.", vbInformation, "¡Aviso!"
        EnfocaControl CmbForPag
        Exit Function
    End If
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Or CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoDeposito Then
        pnMontoPagarCargoDeposito = CDbl(TxtMontoPagar.Text)

        AsignaValorITF
        If Not pbEsMismoTitular Then
            MsgBox "La cuenta de ahorro debe tener el mismo titular de la operación.", vbInformation, "¡Aviso!"
            EnfocaControl txtCuentaCargo
            Exit Function
        End If
        
        If CDbl(TxtMontoPagar.Text) <= 0 Then
            MsgBox "Favor de ingresar el monto de la operación.", vbInformation, "¡Aviso!"
            EnfocaControl TxtMontoPagar
            Exit Function
        End If
        
        
    End If
    ValidaFormaPago = True
End Function

Private Sub IniciarControlesFormaPago()
    CmbForPag.ListIndex = 0
    txtCuentaCargo.NroCuenta = ""
    pnMoneda = 0
    CmbForPag.Enabled = False
    txtCuentaCargo.Visible = False
    lbITFCtaExonerada = False
End Sub

Private Sub IniciarControlesCuentaCargoAbono()
    txtCuentaCargo.NroCuenta = ""
    pnMoneda = 0
'    CmbForPag.Enabled = False
    lbITFCtaExonerada = False
End Sub

Private Sub txtCuentaCargo_KeyPress(KeyAscii As Integer)
    Dim sMsg As String
'    Dim sCta As String
'    sCta = txtCuenta.NroCuenta
    
    If KeyAscii = 13 Then sMsg = ValidaCuentaACargo(txtCuentaCargo.NroCuenta)
    If Len(sMsg) > 0 Then
        MsgBox sMsg, vbInformation, "¡Aviso!"
        EnfocaControl txtCuentaCargo
        Exit Sub
    End If
    Dim sMoneda As String
    pnMoneda = 0
   
    pnMoneda = gMonedaNacional
   
    If Len(txtCuentaCargo.NroCuenta) = 18 Then
        If CInt(Mid(txtCuentaCargo.NroCuenta, 9, 1)) <> pnMoneda Then
            MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la operación de compra/venta.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
    End If
    If lnTipoMov = 1 Then
        ObtieneDatosCuenta txtCuentaCargo.NroCuenta
    Else
        ObtieneDatosCuenta txtCuentaCargo.NroCuenta
    End If
End Sub
Private Function ValidaCuentaACargo(ByVal psCuenta As String) As String
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim sMsg As String
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    sMsg = clsCap.ValidaCuentaOperacion(psCuenta)
    ValidaCuentaACargo = sMsg
End Function

Private Sub ObtieneDatosCuenta(ByVal psCuenta As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsCta As ADODB.Recordset
    Dim rsV As ADODB.Recordset
    Dim lnTpoPrograma As Integer
    Dim lsTieneTarj As String
    Dim lbVistoVal As Boolean
'    Dim lsOpeAhorrCompraVentaCargoCtaAhorro As String
    Dim lsOpeAhorrTransferenciaCargoDeposito As String
    lsOpeAhorrTransferenciaCargoDeposito = ""
    If lnTipoMov = 1 Then
        lsOpeAhorrTransferenciaCargoDeposito = gOpeTransferenciaCargo
    ElseIf lnTipoMov = 2 Then
        lsOpeAhorrTransferenciaCargoDeposito = gOpeTransferenciaDeposito
    End If

    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsV = New ADODB.Recordset
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(psCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
        If sNumTarj = "" Then
            If Mid(psCuenta, 6, 3) = "232" And lnTpoPrograma <> 1 Then
                Set rsV = clsMant.ValidaTarjetizacion(psCuenta, lsTieneTarj)
                If rsV.RecordCount > 0 Then
                    Dim tipoCta As Integer
                    tipoCta = rsCta("nPrdCtaTpo")
                    If tipoCta = 0 Or tipoCta = 2 Then
                        Dim rsCli As New ADODB.Recordset
                        Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
                        Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
                        Dim bExitoSol As Integer
                        Set rsCli = clsCli.GetPersonaCuenta(psCuenta, gCapRelPersTitular)
                        psPersCodTitularAhorroCargoDeposito = rsCli!cperscod ' CTI6
                        psDNITitularAhorroCargoDeposito = rsCli!IDNumero
                        nRespuesta = oSolicitud.SolicitarVistoAtencionSinTarjeta(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(lsOpeAhorrTransferenciaCargoDeposito))
                    
                        If nRespuesta = 1 Then '1:Tiene Visto de atencion sin tarjeta pendiente de autorizar
                             MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                             Exit Sub
                        End If
                        If nRespuesta = 2 Then '2:Tiene visto aceptado
                            MsgBox "La solicitud de atención sin tarjeta fue Aprobada, proceda con la atención", vbInformation, "Aviso"
                        End If
                        If nRespuesta = 3 Then '3:Tiene visto rechazado
                           If MsgBox("La solicitud de atención sin tarjeta fue RECHAZADA. ¿Desea realizar una nueva solicitud?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                Call oSolicitud.RegistrarVistoDeUsuario(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(lsOpeAhorrTransferenciaCargoDeposito))
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        If nRespuesta = 4 Then '4:Se permite registrar la solicitud
                            Dim mensaje As String
                            If lsTieneTarj = "SI" Then
                                mensaje = "El Cliente posee tarjeta. Para continuar deberá registrar el Motivo de Autorización y comunicar al Coordinador o Jefe de Operaciones para su Aprobación. ¿Desea Continuar?"
                            Else
                                mensaje = "El Cliente NO posee tarjeta activa. Para continuar deberá registrar el Motivo de Autorización y comunicar al Coordinador o Jefe de Operaciones para su Aprobación. ¿Desea Continuar?"
                            End If
                        
                            If MsgBox(mensaje, vbInformation + vbYesNo, "Aviso") = vbYes Then
                        
                                bExitoSol = frmCapMotivoAutorizacion.Inicio(gsCodUser, gsCodAge, psCuenta, rsCli!cperscod, CStr(lsOpeAhorrTransferenciaCargoDeposito))
                                If bExitoSol > 0 Then
                                    MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                                End If
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        lbVistoVal = loVistoElectronico.Inicio(5, lsOpeAhorrTransferenciaCargoDeposito)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema"
                            Exit Sub
                        End If
                    End If
                ElseIf lsTieneTarj = "NO" And rsV.RecordCount > 0 Then
                    If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? ", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg ers 065-2017
                        lbVistoVal = loVistoElectronico.Inicio(5, lsOpeAhorrTransferenciaCargoDeposito)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema" 'ADD BY MARG ERS065-2017
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            End If
        Else
            If Mid(psCuenta, 6, 3) = "232" And lnTpoPrograma <> 1 Then
                   Set rsV = clsMant.ValidaTarjetizacion(psCuenta, lsTieneTarj)
                   If rsV.RecordCount > 0 Then
                       Dim tipoCta2 As Integer
                       tipoCta2 = rsCta("nPrdCtaTpo")
                       If tipoCta2 = 0 Or tipoCta = 2 Then
                           Dim rsCli2 As New ADODB.Recordset
                           Dim clsCli2 As New COMNCaptaGenerales.NCOMCaptaGenerales
                           Dim oSolicitud2 As New COMDCaptaGenerales.DCOMCaptaGenerales
                           Dim bExitoSol2 As Integer
                           Dim nRespuesta2 As Integer
                           Set rsCli2 = clsCli2.GetPersonaCuenta(psCuenta, gCapRelPersTitular)
                           psPersCodTitularAhorroCargoDeposito = rsCli2!cperscod
                           psDNITitularAhorroCargoDeposito = rsCli2!IDNumero
                       End If
                   End If
            End If
        End If
        '******************************************************
        Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
        Dim rsPar As ADODB.Recordset
        Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        nProducto = CInt(Mid(psCuenta, 6, 3))
        Set rsPar = clsDef.GetCapParametroNew(nProducto, lnTpoPrograma, psCuenta)
        
        Dim nIdTarifario As Integer
        nIdTarifario = rsPar!nIdTarifario
             
         If nIdTarifario = 1 And EsHaberes(psCuenta) Then
            MsgBox "No puede utilizar esta Operación para una Cuenta de Haberes", vbOKOnly + vbExclamation, App.Title
            CmdGuardar.Enabled = False
            Exit Sub
        End If
        
        Dim dLSCAP As COMDCaptaGenerales.DCOMCaptaGenerales
        Set dLSCAP = New COMDCaptaGenerales.DCOMCaptaGenerales

        If dLSCAP.EsCtaConvenio(psCuenta) Then
            MsgBox "Esta es una cuenta de Convenio." & vbCrLf & "Usar operación Abonos de Ctas de Convenio. ", vbOKOnly + vbInformation, "AVISO"
            Set dLSCAP = Nothing
            Exit Sub
        End If
        '********************************************************
        txtCuentaCargo.Enabled = False
        AsignaValorITF
        CmdGuardar.Enabled = True
        EnfocaControl CmdGuardar
    End If
End Sub

Private Sub AsignaValorITF()
    If Trim(psDNITitularAhorroCargoDeposito) = Trim(txtDNI.Text) Then
        pbEsMismoTitular = True
    Else
        pbEsMismoTitular = False
    End If
    pnITF = 0#
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Or CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoDeposito Then
        If gITF.gbITFAplica Then
            pnITF = Format(gITF.fgITFCalculaImpuesto(pnMontoPagarCargoDeposito), "#0.00")
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(pnITF))
            If nRedondeoITF > 0 Then
                  pnITF = Format(CCur(pnITF) - nRedondeoITF, "#,##0.00")
            End If
         End If
     End If
End Sub

Private Function EsHaberes(ByVal sCta As String) As Boolean
Dim ssql As String
Dim cCap As COMDCaptaGenerales.COMDCaptAutorizacion
Set cCap = New COMDCaptaGenerales.COMDCaptAutorizacion
    EsHaberes = cCap.EsHaberes(sCta)
Set cCap = Nothing
End Function
'
'Private Sub ObtieneDatosCuentaDesembolso(ByVal sCuenta As String)
'Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
'Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
'Dim clsGen As COMDConstSistema.DCOMGeneral
'Dim rsCta As New ADODB.Recordset, rsRel As New ADODB.Recordset
'Dim nEstado As COMDConstantes.CaptacEstado
'Dim nRow As Long
'Dim sMsg As String, sMoneda As String, sPersona As String
'Dim lnTpoPrograma As Integer
'Dim nPersoneria As COMDConstantes.PersPersoneria
'Dim lbVistoVal As Boolean
'
'
'Dim nTipoCuenta As COMDConstantes.ProductoCuentaTipo
'Dim sTipoCuenta As String
'
'
'
'Dim oCapGen As New COMNCaptaGenerales.NCOMCaptaGenerales
'Set oCapGen = New COMNCaptaGenerales.NCOMCaptaGenerales
'
'
'Dim lafirma As frmPersonaFirma
'Dim ClsPersona As COMDPersona.DCOMPersonas
'Dim Rf As ADODB.Recordset
'
'Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
'Dim rsPar As ADODB.Recordset
'Dim nCantOpeCta As Integer
'
'lbVistoVal = False
'
'Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
'sMsg = clsCap.ValidaCuentaOperacion(sCuenta, True)
'Set clsCap = Nothing
'
'If sMsg = "" Then
'    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
'        Set rsCta = New Recordset
'        Set rsCta = clsMant.GetDatosCuenta(sCuenta)
'    Set clsMant = Nothing
'    If Not (rsCta.EOF And rsCta.BOF) Then
'
'        Dim dLSCAP As COMDCaptaGenerales.DCOMCaptaGenerales
'        Set dLSCAP = New COMDCaptaGenerales.DCOMCaptaGenerales
'
'        If dLSCAP.EsCtaConvenio(sCuenta) Then
'            MsgBox "Esta es una cuenta de Convenio." & vbCrLf & "Usar operación Abonos de Ctas de Convenio. ", vbOKOnly + vbInformation, "AVISO"
'            Set dLSCAP = Nothing
'            Exit Sub
'        End If
'
'
'        nEstado = rsCta("nPrdEstado")
'        nPersoneria = rsCta("nPersoneria")
'
'        If nProducto = gCapAhorros Or nProducto = gCapCTS Then
'            lnTpoPrograma = IIf(IsNull(rsCta("nTpoPrograma")), 0, rsCta("nTpoPrograma"))
'        End If
'
'        Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
'        Set rsPar = clsDef.GetCapParametroNew(nProducto, lnTpoPrograma, sCuenta)
'
'        Dim nIdTarifario As Integer
'        nIdTarifario = rsPar!nIdTarifario
'
'         If nIdTarifario = 1 And EsHaberes(sCuenta) And (Trim(nOperacion) <> "200243" And Trim(nOperacion) <> "200244" And Trim(nOperacion) <> "200245") Then
'            MsgBox "No puede utilizar esta Operación para una Cuenta de Haberes", vbOKOnly + vbExclamation, App.Title
'            CmdGuardar.Enabled = False
'            Exit Sub
'        End If
'
'        If nProducto = gCapAhorros Then
'            nParCantDepLib = rsPar!nCantOpeVentDep
'            nParMontoMinDepSol = rsPar!nMontoMinDepSol
'            nParMontoMinDepDol = rsPar!nMontoMinDepDol
'        End If
'        Set rsPar = Nothing
'
'        If bValidaCantDep Then
'            Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
'                nCantOpeCta = clsMant.ObtenerCantidadOperaciones(sCuenta, gCapMovDeposito, gdFecSis)
'            Set clsMant = Nothing
'        End If
'        lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy hh:mm:ss")
'
'        lbITFCtaExonerada = fgITFVerificaExoneracion(sCuenta)
'        fgITFParamAsume Mid(sCuenta, 4, 2), Mid(sCuenta, 6, 3)
'        If nIdTarifario <> 1 And lnTpoPrograma = 6 And (Trim(nOperacion) <> "200243" And Trim(nOperacion) <> "200244" And Trim(nOperacion) <> "200245") Then
'            lbITFCtaExonerada = False
'        End If
'
'
'        Select Case nProducto
'             Case gCapAhorros
'
'             '
'                If rsCta("bOrdPag") Then
'                    lblMensaje = lblMensaje & Chr$(13) & "AHORROS CON ORDEN DE PAGO" & Chr$(13) & sMoneda
'                    pbOrdPag = True
'                Else
'                    'AVMM 10-04-2007
'                    If lnTpoPrograma = 1 Then
'                        lblMensaje = lblMensaje & Chr$(13) & "AHORRO ÑAÑITO" & Chr$(13) & sMoneda
'                    ElseIf lnTpoPrograma = 2 Then
'                        lblMensaje = lblMensaje & Chr$(13) & "AHORROS PANDERITO" & Chr$(13) & sMoneda
'                    ElseIf lnTpoPrograma = 3 Then
'                        '*** PEAC 20090722
'                        'lblMensaje = lblMensaje & Chr$(13) & "AHORROS PANDERO" & Chr$(13) & sMoneda
'                        lblMensaje = lblMensaje & Chr$(13) & "AHORROS POCO A POCO AHORRO" & Chr$(13) & sMoneda
'                    ElseIf lnTpoPrograma = 4 Then
'                        lblMensaje = lblMensaje & Chr$(13) & "AHORROS DESTINO" & Chr$(13) & sMoneda
'                    Else
'                        lblMensaje = lblMensaje & Chr$(13) & "AHORROS SIN ORDEN DE PAGO" & Chr$(13) & sMoneda
'                    End If
'                    pbOrdPag = False
'                End If
'                lblUltContacto = Format$(rsCta("dUltContacto"), "dd mmm yyyy hh:mm:ss")
'
'
'                Set clsGen = Nothing
'        End Select
'
'        lblTipoCuenta = UCase(rsCta("cTipoCuenta"))
'        sTipoCuenta = lblTipoCuenta
'        nTipoCuenta = rsCta("nPrdCtaTpo")
'        lblFirmas = Format$(rsCta("nFirmas"), "#0")
'        Set rsRel = New ADODB.Recordset
'        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
'            Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
'        Set clsMant = Nothing
'        sPersona = ""
'
'        Do While Not rsRel.EOF
'            If rsRel("cPersCod") = gsCodPersUser Then
'                MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
'                Unload Me
'                Exit Sub
'            End If
'        Loop
'
'        rsRel.Close
'        Set rsRel = Nothing
'
'        Dim rsCli As New ADODB.Recordset
'        Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
'        Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
'        Dim bExitoSol As Integer
'        Set rsCli = clsCli.GetPersonaCuenta(psCuenta, gCapRelPersTitular)
'        psPersCodTitularAhorroCargoDeposito = rsCli!cperscod ' CTI6
'        psDNITitularAhorroCargoDeposito = rsCli!IDNumero
'
'
'        txtCuentaCargo.Enabled = False
'        AsignaValorITF
'        CmdGuardar.Enabled = True
'        EnfocaControl CmdGuardar
'    End If
'Else
'    MsgBox sMsg, vbInformation, "Operacion"
'    TxtCuenta.SetFocus
'End If
'End Sub
'***********************************************************************************


