VERSION 5.00
Begin VB.Form frmOpeComisionReprogCred 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comisión: "
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   3120
   ClientWidth     =   5655
   Icon            =   "frmOpeComisionReprogCred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   4320
      TabIndex        =   4
      Top             =   2680
      Width           =   1170
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   3000
      TabIndex        =   3
      Top             =   2680
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   " Cliente: "
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5415
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   495
         Left            =   0
         TabIndex        =   0
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         Texto           =   "Credito"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   330
         Left            =   3720
         TabIndex        =   1
         Top             =   270
         Width           =   375
      End
      Begin VB.Label lblEnvioEstCta 
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
         Left            =   4080
         TabIndex        =   20
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblCtaEnvio 
         Caption         =   "Cuota a enviar:"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   1965
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Prox. Cuota:"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   1250
         Width           =   960
      End
      Begin VB.Label Label10 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   890
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "D.O.I.:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1960
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Subproducto:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1250
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   890
         Width           =   615
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
         Left            =   1200
         TabIndex        =   10
         Top             =   1920
         Width           =   1455
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
         Left            =   1200
         TabIndex        =   9
         Top             =   1560
         Width           =   4095
      End
      Begin VB.Label lblProxCuota 
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
         Left            =   3840
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblSubProducto 
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
         Left            =   1200
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblMoneda 
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
         Left            =   3840
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblMonto 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
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
      Left            =   1320
      TabIndex        =   18
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label lblLabelMonto 
      Caption         =   "Monto S/:"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2760
      Width           =   855
   End
End
Attribute VB_Name = "frmOpeComisionReprogCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmOpeComisionReprogCred
'** Descripción : Formulario para registrar la comision por reprogramacion de credito creado segun TI-ERS029-2013
'** Creación : JUEZ, 20130411 09:00:00 AM
'**********************************************************************************************

Option Explicit

'Dim lsOpeCod As Long
Dim fsOpeCod As Long 'JUEZ 20130529
Dim lsPrdConceptoCod As Integer
Dim fsPersCod As String

'Private Sub Form_Load()
'    fsOpeCod = gComisionReprogCredito
'    ActXCodCta.CMAC = gsCodCMAC
'    ActXCodCta.Age = gsCodAge
'    cmdAceptar.Enabled = False
'End Sub
Public Sub Inicia(ByVal psOpeCod As CaptacOperacion, ByVal psTitulo As String)
    fsOpeCod = psOpeCod
    ActXCodCta.CMAC = gsCodCMAC
    ActXCodCta.Age = gsCodAge
    Me.Caption = Me.Caption & psTitulo
    cmdAceptar.Enabled = False
    'JUEZ 20130529 *******************************
    'If fsOpeCod = gComisionEnvioEstadoCta Then
    If fsOpeCod = gComiCredEstadoCta Then 'JUEZ 20150928
        lblCtaEnvio.Visible = True
        lblEnvioEstCta.Caption = ""
        lblEnvioEstCta.Visible = True
    Else
        lblCtaEnvio.Visible = False
        lblEnvioEstCta.Caption = ""
        lblEnvioEstCta.Visible = False
    End If
    'END JUEZ ************************************
    Me.Show 1
End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If Len(ActXCodCta.NroCuenta) = 18 Then
        Dim oCred As COMDCredito.DCOMCredito
        Set oCred = New COMDCredito.DCOMCredito
        'JUEZ 20130529 ******************************************************************
        'If fsOpeCod = gComisionReprogCredito Then
        If fsOpeCod = gComiCredReprogCred Then 'JUEZ 20150928
            If oCred.ExisteComisionVigente(ActXCodCta.NroCuenta, fsOpeCod) = False Then
                CargarDatos
                ActXCodCta.Enabled = False 'JUEZ 20151229
                cmdBuscar.Enabled = False 'JUEZ 20151229
                cmdAceptar.Enabled = True
                cmdAceptar.SetFocus
            Else
                MsgBox "Ya existe un pago vigente de la comisión por reprogramación de este crédito", vbInformation, "Aviso"
            End If
        Else
            CargarDatos
        End If
        'END JUEZ ***********************************************************************
    End If
End Sub

Private Sub cmdAceptar_Click()

    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII

    If MsgBox("Desea Grabar la Información?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Dim oCred As COMDCredito.DCOMCredActBD
    Set oCred = New COMDCredito.DCOMCredActBD
    Dim oCredMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oCredMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCont As COMNContabilidad.NCOMContFunciones
    Set clsCont = New COMNContabilidad.NCOMContFunciones
    Dim lsMov As String
    Dim lsBoleta As String
    'JUEZ 20130529 *********************************************
    Dim lsGlosa As String
    Dim lsTitVoucher As String
    
    'If fsOpeCod = gComisionReprogCredito Then
    If fsOpeCod = gComiCredReprogCred Then 'JUEZ 20150928
        lsGlosa = "Comision por Reprogramacion de Credito"
        lsTitVoucher = "REPROGRAMACION DE CREDITO"
    'ElseIf fsOpeCod = gComisionEnvioEstadoCta Then
    ElseIf fsOpeCod = gComiCredEstadoCta Then 'JUEZ 20150928
        lsGlosa = "Comision por Envio Estado Cuenta - Cuota " & lblEnvioEstCta.Caption & " - "
        lsTitVoucher = "ENVIO ESTADO CUENTA"
    End If
    'END JUEZ **************************************************
    
    lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    gnMovNro = 0
    Call oCredMov.OtrasOperaciones(lsMov, fsOpeCod, CDbl(lblComision.Caption), ActXCodCta.NroCuenta, lsGlosa & " cliente: " & lblCliente.Caption, gMonedaNacional, fsPersCod, , , , , , , gnMovNro)
    If gnMovNro <> 0 Then
        'Call oCred.dInsertComision(gnMovNro, ActXCodCta.NroCuenta, CDbl(lblComision.Caption), IIf(fsOpeCod = gComisionEnvioEstadoCta, CInt(IIf(lblEnvioEstCta.Caption = "", 0, lblEnvioEstCta.Caption)), 0))
        Call oCred.dInsertComision(gnMovNro, ActXCodCta.NroCuenta, CDbl(lblComision.Caption), IIf(fsOpeCod = gComiCredEstadoCta, CInt(IIf(lblEnvioEstCta.Caption = "", 0, lblEnvioEstCta.Caption)), 0)) 'JUEZ 20150928
        Set oCred = Nothing
        Dim oBol As COMNCredito.NCOMCredDoc
        Set oBol = New COMNCredito.NCOMCredDoc
            lsBoleta = oBol.ImprimeBoletaComision(lsTitVoucher, Left("Total pago comision", 36), "", Str(CDbl(lblComision.Caption)), lblCliente.Caption, lblDOI.Caption, "________" & Mid(ActXCodCta.NroCuenta, 9, 1), False, "", "", , gdFecSis, gsNomAge, gsCodUser, sLpt, , gbImpTMU)
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
        Limpiar
        'INICIO JHCU ENCUESTA 16-10-2019
        Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
        'FIN
    Else
        MsgBox "Hubo un error en el registro", vbInformation, "Aviso"
    End If
End Sub

Private Sub Limpiar()
    ActXCodCta.Prod = ""
    ActXCodCta.Cuenta = ""
    lblMonto.Caption = ""
    LblMoneda.Caption = ""
    lblSubProducto.Caption = ""
    lblProxCuota.Caption = ""
    lblCliente.Caption = ""
    lblDOI.Caption = ""
    lblComision.Caption = ""
    lblEnvioEstCta.Caption = ""
    cmdAceptar.Enabled = False
    ActXCodCta.Enabled = True 'JUEZ 20151229
    cmdBuscar.Enabled = True 'JUEZ 20151229
End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona
    Limpiar
    Set oPers = frmBuscaPersona.inicio()
    If Not oPers Is Nothing Then
        Call FrmVerCredito.inicio(oPers.sPersCod, , , True, frmOpeComisionReprogCred.ActXCodCta)
        ActXCodCta.SetFocusCuenta
    End If
    Set oPers = Nothing
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    ActXCodCta.SetFocus 'JUEZ 20151229
End Sub

Private Sub CargarDatos()
    Dim oCred As COMDCredito.DCOMCredito
    Dim R As ADODB.Recordset
    Dim nCuotaNueva As Integer
    Dim nValorCom As Double 'JUEZ 20151229
    Dim nTCVenta As Double 'JUEZ 20151229
    
    'JUEZ 20151229 *******************************************
    Dim oDGeneral As COMDConstSistema.NCOMTipoCambio
    Set oDGeneral = New COMDConstSistema.NCOMTipoCambio
        nTCVenta = oDGeneral.EmiteTipoCambio(gdFecSis, TCVenta)
    Set oDGeneral = Nothing
    'END JUEZ ************************************************
    
    Set oCred = New COMDCredito.DCOMCredito
    Set R = oCred.RecuperaDatosComision(ActXCodCta.NroCuenta, 1)
    If Not (R.EOF And R.BOF) Then 'WIOR 20130713
        'JUEZ 20130529 ********************************************************************
        'lsPrdConceptoCod = IIf(R!nPersoneria = 1, gColocConceptoCodGastoComisionReprogNat, gColocConceptoCodGastoComisionReprogJur)
        'If fsOpeCod = gComisionReprogCredito Then
        If fsOpeCod = gComiCredReprogCred Then 'JUEZ 20150928
            lsPrdConceptoCod = IIf(R!nPersoneria = 1, gColocConceptoCodGastoComisionReprogNat, gColocConceptoCodGastoComisionReprogJur)
        'ElseIf fsOpeCod = gComisionEnvioEstadoCta Then
        ElseIf fsOpeCod = gComiCredEstadoCta Then 'JUEZ 20150928
            If oCred.ExisteComisionEnvioEstadoCtaCalendario(ActXCodCta.NroCuenta, R!nNroProxCuota) Then
                MsgBox "La cuota " & R!nNroProxCuota & " del crédito ya tiene la comisión por envio de estado de cuenta en el Plan de Pagos", vbInformation, "Aviso"
                cmdAceptar.Enabled = False
                Exit Sub
            ElseIf oCred.ExisteComisionEnvioEstadoCta(ActXCodCta.NroCuenta, R!nNroProxCuota) Then
                If oCred.ExisteCuotaCalendario(ActXCodCta.NroCuenta, R!nNroProxCuota + 1) = False Then
                    MsgBox "No se puede pagar mas comisiones porque no hay cuotas pendientes de pago por envio de estado de cuenta", vbInformation, "Aviso"
                    cmdAceptar.Enabled = False
                    Exit Sub
                Else
                    nCuotaNueva = R!nNroProxCuota + 1
                End If
            Else
                nCuotaNueva = R!nNroProxCuota
            End If
            lsPrdConceptoCod = gColocConceptoCodGastoComisionEnvioEstCta
        End If
        'END JUEZ *************************************************************************
        lblMonto.Caption = Format(R!nSaldo, "#,##0.00")
        LblMoneda.Caption = R!cmoneda
        lblSubProducto.Caption = R!cTpoProdDesc
        lblProxCuota.Caption = Format(R!dProxCuota, "dd/mm/yyyy")
        lblCliente.Caption = R!cPersNombre
        lblDOI.Caption = R!cPersIDnro
        fsPersCod = R!cPersCod
        'JUEZ 20130529 ****************************
        'If fsOpeCod = gComisionEnvioEstadoCta Then
        If fsOpeCod = gComiCredEstadoCta Then 'JUEZ 20150928
            lblEnvioEstCta = nCuotaNueva
        End If
        'END JUEZ *********************************
        
        Set R = oCred.RecuperaProductoConcepto(lsPrdConceptoCod)
        Set oCred = Nothing
        'JUEZ 20151229 **************************************
        'lblComision.Caption = Format(R!nValor, "#,##0.00")
        nValorCom = R!nValor
        If fsOpeCod = gComiCredReprogCred Then
            nValorCom = Format(nValorCom / IIf(Mid(ActXCodCta.NroCuenta, 9, 1) = gMonedaNacional, 1, nTCVenta), "#,##0.00")
            lblLabelMonto.Caption = "Monto " & IIf(Mid(ActXCodCta.NroCuenta, 9, 1) = gMonedaNacional, "S/", "$") & ":"
        End If
        lblComision.Caption = Format(nValorCom, "#,##0.00")
        'END JUEZ *******************************************
        Set R = Nothing
        cmdAceptar.Enabled = True
        cmdAceptar.SetFocus
    'WIOR 20130713 *******************
    Else
        MsgBox "No existen Datos.", vbInformation, "Aviso"
        Limpiar
    End If
    'WIOR FIN ************************
End Sub
