VERSION 5.00
Begin VB.Form frmOpeComisionOtros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comisión: "
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   Icon            =   "frmOpeComisionOtros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFormaPago 
      Height          =   800
      Left            =   120
      TabIndex        =   10
      Top             =   1900
      Width           =   6700
      Begin VB.ComboBox CmbForPag 
         Height          =   315
         Left            =   1100
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   250
         Width           =   1800
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   495
         Left            =   3000
         TabIndex        =   15
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   873
         Texto           =   "Cuenta Nº :"
      End
      Begin VB.Label lblFormaPago 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   250
         Width           =   855
      End
      Begin VB.Label LblNumDoc 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4300
         TabIndex        =   13
         Top             =   250
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblNroDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   3105
         TabIndex        =   12
         Top             =   250
         Visible         =   0   'False
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4320
      TabIndex        =   5
      Top             =   2760
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   5640
      TabIndex        =   4
      Top             =   2760
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6700
      Begin SICMACT.TxtBuscar TxtBCodPers 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
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
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1100
         TabIndex        =   9
         Top             =   720
         Width           =   5500
      End
      Begin VB.Label lblDOI 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1100
         TabIndex        =   8
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Label Label2 
         Caption         =   "D.O.I. : "
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1245
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente: "
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Width           =   615
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Monto S/:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2850
      Width           =   840
   End
   Begin VB.Label lblComision 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
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
      Height          =   300
      Left            =   1220
      TabIndex        =   6
      Top             =   2800
      Width           =   1800
   End
End
Attribute VB_Name = "frmOpeComisionOtros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmOpeComisionOtros
'** Descripción : Formulario para registrar la comision para otros conceptos creado segun TI-ERS029-2013
'** Creación : JUEZ, 20130411 09:00:00 AM
'**********************************************************************************************

Option Explicit
Dim R As ADODB.Recordset
Dim fsOpeCod As Long
Dim fsConceptoCod As Integer
Dim fsGlosa As String
Dim fsTitVoucher As String
Private nMontoVoucher As Currency 'CTI5 ERS0112020
Dim sNumTarj As String 'CTI5 ERS0112020
Dim loVistoElectronico As frmVistoElectronico
Dim nRespuesta As Integer 'CTI4 ERS0112020

Public Sub Inicia(ByVal psOpeCod As Long, ByVal pnConcepto As Integer, ByVal psTitulo As String, ByVal psGlosa As String, ByVal psTitVoucher As String)
    Dim loParam As COMDColocPig.DCOMColPCalculos
    fsOpeCod = psOpeCod
    fsConceptoCod = pnConcepto
    fsGlosa = psGlosa
    fsTitVoucher = psTitVoucher
    
    Me.Caption = Me.Caption & psTitulo
    
    Set loParam = New COMDColocPig.DCOMColPCalculos
    lblComision.Caption = Format(loParam.dObtieneColocParametro(fsConceptoCod), "#,##0.00")
    Set loParam = Nothing
    Me.Show 1
End Sub

Private Sub cmdAceptar_Click()
    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    Dim lsNombreClienteCargoCta As String 'CTI5 ERS0112020
    Dim loGrabarCan As COMNColoCPig.NCOMColPContrato 'CTI5 ERS0112020
    Set loGrabarCan = New COMNColoCPig.NCOMColPContrato 'CTI5 ERS0112020
    Dim MatDatosAho(14) As String 'CTI5 ERS0112020
    Dim lsBoletaAhorro As String 'CTI5 ERS0112020

    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII
    
    If Not ValidaFormaPago Then Exit Sub 'CTI5 ERS0112020
    
    If MsgBox("Desea Grabar la Información?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Dim oCredMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oCredMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCont As COMNContabilidad.NCOMContFunciones
    Set clsCont = New COMNContabilidad.NCOMContFunciones
    Dim lsMov As String
    Dim lsBoleta As String

    'CTI5 ERS0112020
    Select Case CInt(Trim(Right(CmbForPag.Text, 10)))
        Case gColocTipoPagoEfectivo
            fsOpeCod = gComiCredConstNoAdeudo
        Case gColocTipoPagoVoucher
            fsOpeCod = gColPOpeCancelVoucher
        Case gColocTipoPagoCargoCta
            fsOpeCod = gComiCredConstNoAdeudoCargoCta
    End Select
    
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then lsNombreClienteCargoCta = PstaNombre(loGrabarCan.ObtieneNombreTitularCargoCta(AXCodCta.NroCuenta))
    'END CTI5 ERS0112020
        
    lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    gnMovNro = 0
    
    Dim lsFechaHoraGrab As String               'CTI5 ERS0112020
    lsFechaHoraGrab = fgFechaHoraGrab(lsMov)    'CTI5 ERS0112020
    
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
        'gnMovNro = oCredMov.OtrasOperaciones(lsMov, fsOpeCod, CDbl(lblComision.Caption), "", fsGlosa & " cliente: " & lblCliente.Caption, gMonedaNacional, TxtBCodPers.Text, , , , , , , gnMovNro) 'CTI5 ERS0112020
        gnMovNro = oCredMov.OtrasOperaciones(lsMov, fsOpeCod, CDbl(lblComision.Caption), "", fsGlosa & " cliente: " & lblCliente.Caption, gMonedaNacional, TxtBCodPers.Text, , , , , , , , , , , AXCodCta.NroCuenta, MatDatosAho, gColocTipoPagoCargoCta, lsFechaHoraGrab) 'CTI5 ERS0112020
    Else
        gnMovNro = oCredMov.OtrasOperaciones(lsMov, fsOpeCod, CDbl(lblComision.Caption), "", fsGlosa & " cliente: " & lblCliente.Caption, gMonedaNacional, TxtBCodPers.Text)
    End If
    
    If gnMovNro <> 0 Then
    
        'If fsOpeCod = gComisionEvalPolEndosada Then
        Select Case fsOpeCod 'JUEZ 20150928
            Case gComiCredEvalPolizaEnd, gComiCredConstNoAdeudo
                Dim oCred As COMDCredito.DCOMCredActBD
                
                Set oCred = New COMDCredito.DCOMCredActBD
                Call oCred.dInsertComision(gnMovNro, TxtBCodPers.Text, CDbl(lblComision.Caption))
                Set oCred = Nothing
        End Select
        
        Dim oBol As COMNCredito.NCOMCredDoc
        Set oBol = New COMNCredito.NCOMCredDoc
        
'        'CTI5 ERS0112020
'        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
'            Set loGrabarCan = New COMNColoCPig.NCOMColPContrato
'            lsNombreClienteCargoCta = PstaNombre(loGrabarCan.ObtieneNombreTitularCargoCta(AXCodCta.NroCuenta)) 'CTI5 ERS0112020
'            lsBoletaAhorro = oBol.ImprimeBoletaAhorro("RETIRO AHORROS", "CARGO CTA. - CONSTANCIA NO ADEUDO", fsOpeCod, CStr(lblComision.Caption), lsNombreClienteCargoCta, AXCodCta.NroCuenta, "", 0, "", "", 1, 0, , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt, gsCodCMAC, 0)
'        End If
'        'END CTI5 ERS0112020
        lsBoleta = oBol.ImprimeBoletaComision(fsTitVoucher, Left("Total pago comision", 36), "", Str(CDbl(lblComision.Caption)), lblCliente.Caption, lblDOI.Caption, "________" & gMonedaNacional, False, "", "", , gdFecSis, gsNomAge, gsCodUser, sLpt, , gbImpTMU, , , , IIf(CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta, AXCodCta.NroCuenta, ""))
        lsBoleta = lsBoleta  'CTI5 ERS0112020
        Set oBol = Nothing
        
        Do
           If Trim(lsBoleta) <> "" Then
                lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    Print #nFicSal, ""
                Close #nFicSal
          End If
            
        Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
        Set oBol = Nothing
        
        'CTI4 ERS0112020
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            Dim oMovOperacion As COMDMov.DCOMMov
            Dim nMovNroOperacion As Long
            Dim rsCli As New ADODB.Recordset
            Dim clsCli As New COMNCaptaGenerales.NCOMCaptaGenerales
            Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
            Set oMovOperacion = New COMDMov.DCOMMov
            nMovNroOperacion = oMovOperacion.GetnMovNro(lsMov)

            loVistoElectronico.RegistraVistoElectronico nMovNroOperacion, , gsCodUser, nMovNroOperacion

            If nRespuesta = 2 Then
                Set rsCli = clsCli.GetPersonaCuenta(AXCodCta.NroCuenta, gCapRelPersTitular)
                oSolicitud.ActualizarCapAutSinTarjetaVisto_nMovNro gsCodUser, gsCodAge, AXCodCta.NroCuenta, rsCli!cPersCod, nMovNroOperacion, CStr(gCargoCtaComiCredConstNoAdeudo)
            End If
            Set oMovOperacion = Nothing
            nRespuesta = 0
        End If
        'CTI4 end
        
        Limpiar
        'INICIO JHCU ENCUESTA 16-10-2019
        Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
        'FIN
    Else
        MsgBox "Hubo un error en el registro", vbInformation, "Aviso"
    End If
End Sub

Private Sub Limpiar()
    lblCliente.Caption = ""
    lblDOI.Caption = ""
    TxtBCodPers.Text = ""
    cmdAceptar.Enabled = False
    sNumTarj = "" 'CTI4 ERS0112020
    AXCodCta.NroCuenta = ""
    CmbForPag.ListIndex = -1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub TxtBCodPers_EmiteDatos()
    If TxtBCodPers.Text <> "" Then
        Dim oCred As COMDCredito.DCOMCredito
        Set oCred = New COMDCredito.DCOMCredito
        'If fsOpeCod = gComisionEvalPolEndosada Then
        If fsOpeCod = gComiCredEvalPolizaEnd Then 'JUEZ 20150928
            If oCred.ExisteCreditoTitular(TxtBCodPers.Text, True, True, , True) Then 'WIOR 20130829 MOSTRAR CREDITOS VIGENTES
                CargaDatos
            Else
                MsgBox "Sólo se puede realizar pago por este concepto a nombre de un titular de crédito en estado aprobado", vbInformation, "Aviso"
                Limpiar
            End If
        ElseIf fsOpeCod = gComisionDupTasacion Then
            If oCred.ExisteCreditoTitular(TxtBCodPers.Text, , True, True) Then
                CargaDatos
            Else
                MsgBox "Sólo se puede realizar pago por este concepto a nombre de un titular de crédito", vbInformation, "Aviso"
                Limpiar
            End If
        'JUEZ 20151229 ****************************************
        ElseIf fsOpeCod = gComiCredConstNoAdeudo Then
            If ValidarPrimeraConstNoAdeudo Then
                CargaDatos
                'MsgBox "No es requerido el pago de comisión para el cliente seleccionado por ser su primera solicitud", vbInformation, "Aviso"
                MsgBox "Cliente no realiza el pago, debido a que la primera constancia es Gratis.", vbInformation, "Aviso" 'APRI20180620 ERS004-2018
                Limpiar
            Else
                If Not oCred.ExisteSolicitudPendiente(TxtBCodPers.Text) Then 'APRI20180620 ERS004-2018
                    CargaDatos
                    MsgBox "Cliente no cuenta con Solicitud para realizar pago.", vbInformation, "Aviso"
                    Limpiar
                ElseIf oCred.ExisteComisionVigente(TxtBCodPers.Text, gComiCredConstNoAdeudo) Then
                    CargaDatos
                    MsgBox "Ya existe un pago de comisión vigente del cliente seleccionado", vbInformation, "Aviso"
                    Limpiar
                Else
                    CargaDatos
                End If
            End If
        'END JUEZ *********************************************
        Else
            CargaDatos
        End If
    Else
        Limpiar
    End If
    Set oCred = Nothing
End Sub

Private Sub CargaDatos()
    Dim oCred As COMDCredito.DCOMCredito
    Set oCred = New COMDCredito.DCOMCredito
    Set R = oCred.RecuperaDatosComision(TxtBCodPers.Text, 2)
    Set oCred = Nothing
    lblCliente.Caption = R!cPersNombre
    lblDOI.Caption = R!cPersIDnro
    Set R = Nothing
    cmdAceptar.Enabled = True
End Sub

'JUEZ 20151229 ***************************************
Private Function ValidarPrimeraConstNoAdeudo() As Boolean
    Dim objCOMNCredito As COMNCredito.NCOMCredito
    Set objCOMNCredito = New COMNCredito.NCOMCredito
    Dim rs As ADODB.Recordset
    
    Set rs = objCOMNCredito.ValidarConstNoAdeudoXPersona(TxtBCodPers.Text)
    'If rs.RecordCount = 0 Then
    If Not rs.EOF And Not rs.BOF Then
        If rs!nTotal = 1 And rs!nPendiente = 1 Then 'APRI20180620 ERS004-2017
            ValidarPrimeraConstNoAdeudo = True
        Else
            ValidarPrimeraConstNoAdeudo = False
        End If
    'Else
        'ValidarPrimeraConstNoAdeudo = False
    End If
End Function
'END JUEZ ********************************************

'CTI5 ERS0112020 *****************
Private Sub CmbForPag_Click()
    EstadoFormaPago IIf(CmbForPag.ListIndex = -1, -1, CInt(Trim(Right(IIf(CmbForPag.Text = "", "-1", CmbForPag.Text), 10))))
    If CmbForPag.ListIndex <> -1 Then
        If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher Then
     
            Dim oformVou As New frmCapRegVouDepBus
            Dim lnTipMot As Integer
            Dim sGlosa As String
            Dim sIF As String
            Dim sVaucher As String
            Dim sPersCod As String
            Dim sNombre As String
            Dim sDireccion As String
            Dim sDocumento As String
                        
            lnTipMot = 15 ' Cancelacion Credito Pignoraticio
            'oformVou.iniciarFormularioDeposito CInt(Mid(AXCodCta.NroCuenta, 9, 1)), lnTipMot, sGlosa, sIF, sVaucher, nMontoVoucher, sPersCod, nMovNroRVD, nMovNroRVDPend, sNombre, sDireccion, sDocumento, AXCodCta.NroCuenta
            If Len(sVaucher) = 0 Then Exit Sub
            LblNumDoc.Caption = Trim(Mid(sVaucher, 1, Len(sVaucher) - 10))
            'cmdGrabar.Enabled = True
        ElseIf CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
            Dim sCuenta As String
            
            sCuenta = frmATMCargaCuentas.RecuperaCuenta(CStr(gComiCredConstNoAdeudoCargoCta), sNumTarj, 232, False)
            If Val(Mid(sCuenta, 6, 3)) <> "232" And sCuenta <> "" Then
                MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
                Exit Sub
            End If
            If sCuenta <> "" Then
                AXCodCta.NroCuenta = sCuenta
                AXCodCta.SetFocusCuenta
            End If
            If Len(sCuenta) = 18 Then
                If CInt(Mid(sCuenta, 9, 1)) <> CInt(Mid(AXCodCta.NroCuenta, 9, 1)) Then
                    MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la cuenta a amortizar.", vbOKOnly + vbInformation, App.Title
                End If
            End If
            If Len(sCuenta) = 0 Then
                AXCodCta.EnabledAge = True
                AXCodCta.EnabledCta = True
                AXCodCta.SetFocusAge
                Exit Sub
            End If
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.Enabled = False
            'AsignaValorITF
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
        End If
    End If
End Sub
Private Sub EstadoFormaPago(ByVal nFormaPago As Integer)
    LblNumDoc.Caption = ""
    AXCodCta.NroCuenta = ""
    Select Case nFormaPago
        Case -1
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            AXCodCta.Visible = False
            cmdAceptar.Enabled = True
        Case gColocTipoPagoEfectivo
            AXCodCta.Visible = False
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            lblNroDocumento.Visible = False
            cmdAceptar.Enabled = True
        Case gColocTipoPagoCargoCta
            LblNumDoc.Visible = False
            lblNroDocumento.Visible = False
            AXCodCta.Visible = True
            AXCodCta.Enabled = True
            AXCodCta.CMAC = gsCodCMAC
            AXCodCta.Prod = Trim(Str(gCapAhorros))
            'AXCodCta.NroCuenta = sCuenta '---
            cmdAceptar.Enabled = False
        Case gColocTipoPagoVoucher
            LblNumDoc.Visible = True
            lblNroDocumento.Visible = True
            AXCodCta.Visible = False
            cmdAceptar.Enabled = False
    End Select
End Sub
Private Function ValidaFormaPago() As Boolean
Dim clsCap As New COMNCaptaGenerales.NCOMCaptaMovimiento
    ValidaFormaPago = False
    If CmbForPag.ListIndex = -1 Then
        MsgBox "No se ha seleccionado la forma de pago. Verifique.", vbInformation, "¡Aviso!"
        EnfocaControl CmbForPag
        Exit Function
    End If
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher And Len(Trim(LblNumDoc.Caption)) = 0 Then
        MsgBox "No se ha seleccionado el voucher correctamente. Verifique.", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoVoucher And Len(Trim(LblNumDoc.Caption)) > 0 _
        And CCur(lblComision.Caption) > CCur(nMontoVoucher) Then
        MsgBox "No se puede realizar el Pago con Voucher solo dispone de: " & Format(nMontoVoucher, "#0.00") & ". Verifique.", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
    
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta And Len(AXCodCta.NroCuenta) <> 18 Then
        MsgBox "Favor de ingresar la cuenta de ahorros a debitar", vbInformation, "¡Aviso!"
        CmbForPag.SetFocus
        Exit Function
    End If
        
    If CInt(Trim(Right(CmbForPag.Text, 10))) = gColocTipoPagoCargoCta Then
        If Not clsCap.ValidaSaldoCuenta(AXCodCta.NroCuenta, CDbl(lblComision.Caption)) Then
            MsgBox "Cuenta a debitar NO posee saldo suficiente o NO está ACTIVA", vbInformation, "¡Aviso!"
            Exit Function
        End If
    End If
    
    ValidaFormaPago = True
End Function
Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    Dim sMsg As String
    If KeyAscii = 13 Then sMsg = ValidaCuentaACargo(AXCodCta.NroCuenta)
    If Len(sMsg) > 0 Then
        MsgBox sMsg, vbInformation, "¡Aviso!"
        AXCodCta.SetFocus
        Exit Sub
    End If
    If Len(AXCodCta.NroCuenta) = 18 Then
        If CInt(Mid(AXCodCta.NroCuenta, 9, 1)) <> CInt(Mid(AXCodCta.NroCuenta, 9, 1)) Then
            MsgBox "La cuenta de ahorro no tiene el mismo tipo de moneda que la cuenta a amortizar.", vbOKOnly + vbInformation, App.Title
        End If
    End If
    ObtieneDatosCuenta AXCodCta.NroCuenta
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
                        nRespuesta = oSolicitud.SolicitarVistoAtencionSinTarjeta(gsCodUser, gsCodAge, psCuenta, rsCli!cPersCod, CStr(gCargoCtaComiCredConstNoAdeudo))
                    
                        If nRespuesta = 1 Then '1:Tiene Visto de atencion sin tarjeta pendiente de autorizar
                             MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                             Exit Sub
                        End If
                        If nRespuesta = 2 Then '2:Tiene visto aceptado
                            MsgBox "La solicitud de atención sin tarjeta fue Aprobada, proceda con la atención", vbInformation, "Aviso"
                        End If
                        If nRespuesta = 3 Then '3:Tiene visto rechazado
                           If MsgBox("La solicitud de atención sin tarjeta fue RECHAZADA. ¿Desea realizar una nueva solicitud?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                                Call oSolicitud.RegistrarVistoDeUsuario(gsCodUser, gsCodAge, psCuenta, rsCli!cPersCod, CStr(gCargoCtaComiCredConstNoAdeudo))
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
                        
                                bExitoSol = frmCapMotivoAutorizacion.Inicio(gsCodUser, gsCodAge, psCuenta, rsCli!cPersCod, CStr(gCargoCtaComiCredConstNoAdeudo))
                                If bExitoSol > 0 Then
                                    MsgBox "La solicitud de atención sin tarjeta fue enviada. " & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
                                End If
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        End If
                        lbVistoVal = loVistoElectronico.Inicio(5, gCargoCtaComiCredConstNoAdeudo)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema"
                            Exit Sub
                        End If
                    End If
                ElseIf lsTieneTarj = "NO" And rsV.RecordCount > 0 Then
                    If MsgBox("El Cliente debe solicitar su tarjeta para realizar las operaciones, si desea continuar con la operacion? ", vbInformation + vbYesNo, "Mensaje del Sistema") = vbYes Then 'add by marg ers 065-2017
                        lbVistoVal = loVistoElectronico.Inicio(5, gCargoCtaComiCredConstNoAdeudo)
                        If Not lbVistoVal Then
                            MsgBox "Visto Incorrecto por favor comunicar al supervisor de operaciones", vbInformation, "Mensaje del Sistema" 'ADD BY MARG ERS065-2017
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            End If
        End If
        AXCodCta.Enabled = False
        'AsignaValorITF
        cmdAceptar.Enabled = True
        cmdAceptar.SetFocus
    End If
End Sub
Private Sub Form_Load()
    Call CargaControles
    
    If fsOpeCod = 322004 Then
        fraFormaPago.Visible = False
        CmbForPag.ListIndex = 0
    Else
        fraFormaPago.Visible = True
    End If

End Sub
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oCons = New COMDConstantes.DCOMConstantes
    Set R = oCons.RecuperaConstantes(gColocTipoPago, , , 3)
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(R, CmbForPag)
    Set loVistoElectronico = New frmVistoElectronico
    Exit Sub
ERRORCargaControles:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub
'END CTI5
