VERSION 5.00
Begin VB.Form frmCredPagoTransferidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago de Creditos Transferidos - FOCMAC"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   Icon            =   "frmCredPagoTransferidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   120
      TabIndex        =   36
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      TabIndex        =   35
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5640
      TabIndex        =   34
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "Montos"
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
      Left            =   4320
      TabIndex        =   21
      Top             =   3240
      Width           =   2415
      Begin SICMACT.EditMoney AXMontoPago 
         Height          =   285
         Left            =   1200
         TabIndex        =   28
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin SICMACT.EditMoney TxtTotalAPagar 
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Total Pagar:"
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
         Left            =   120
         TabIndex        =   33
         Top             =   990
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "I.T.F."
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   31
         Top             =   645
         Width           =   375
      End
      Begin VB.Label LblITF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   285
         Left            =   1200
         TabIndex        =   30
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pagos"
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
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   4095
      Begin VB.TextBox txtGlosa 
         Height          =   300
         Left            =   795
         TabIndex        =   26
         Top             =   960
         Width           =   3135
      End
      Begin VB.ComboBox cboModalidad 
         Height          =   315
         ItemData        =   "frmCredPagoTransferidos.frx":030A
         Left            =   795
         List            =   "frmCredPagoTransferidos.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   240
         Width           =   1950
      End
      Begin VB.TextBox txtNumDoc 
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
         Height          =   300
         Left            =   795
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Label Label12 
         Caption         =   "Glosa:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   285
         Width           =   360
      End
      Begin VB.Label LblNumDoc 
         AutoSize        =   -1  'True
         Caption         =   "Nº Doc:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   570
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Saldos"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   6615
      Begin VB.Label Label11 
         Caption         =   "Gastos:"
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Interés:"
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Mora:"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Capital:"
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   285
         Width           =   615
      End
      Begin VB.Label lblGasto 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4320
         TabIndex        =   15
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label lblInteres 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4320
         TabIndex        =   14
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lblMora 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1410
         TabIndex        =   13
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label lblCapital 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1410
         TabIndex        =   12
         Top             =   240
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Crédito"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6615
      Begin VB.Label lblMoneda 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label2 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "DOI:"
         Height          =   255
         Left            =   420
         TabIndex        =   8
         Top             =   750
         Width           =   375
      End
      Begin VB.Label lblDoi 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   930
         TabIndex        =   7
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label lblMetLiquid 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5400
         TabIndex        =   6
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Met.Liquid:"
         Height          =   195
         Index           =   4
         Left            =   4560
         TabIndex        =   5
         Top             =   750
         Width           =   780
      End
      Begin VB.Label lblCliente 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   930
         TabIndex        =   4
         Top             =   360
         Width           =   5265
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   525
      End
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "&Buscar..."
      Height          =   360
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   1020
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   820
      Texto           =   "Crédito"
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmCredPagoTransferidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'RIRO20140530 ERS017 ********************
Option Explicit
Dim nMovNroRVD As Long
Dim nMovNroRVDPen As Long
Dim nMontoVoucher As Currency
'END RIRO *******************************
Dim fnVarOpeCod As Long
Dim fnSaldoCap As Currency, fnSaldoIntComp As Currency, fnSaldoIntMorat As Currency, fnSaldoGasto As Currency
Dim fnCapPag As Currency, fnIntCompPag As Currency, fnIntMoratPag As Currency, fnGastoPag As Currency
Dim fnNewSaldoCap As Currency, fnNewSaldoIntComp As Currency, fnNewSaldoIntMorat As Currency, fnNewSaldoGasto As Currency
Dim fmMatGastos As Variant ' 1=nNroGastoCta//2=nMonto//3=nMontoPagado//4=nColocRecGastoEstado//5=Modificado
Dim fnNroCalend As Integer, fnNroUltGastoCta As Integer
Dim fsVarPersCodCMAC As String
Dim fnRegCancelacion As Integer
Dim fnMontoPagar As Currency
Dim fnEstadoNew As ColocEstado
Dim fnEstadoIni As ColocEstado
Dim fnEstadoTransf As ColocEstado
Dim fnTotalDeuda As Currency
Dim nRedondeoITF As Double
Dim oDocRec As UDocRec
'**Datos para la distribución **********************************************
Dim fnCapDist As Currency, fnIntCompDist As Currency, fnIntMoratDist As Currency, fnGastoDist As Currency
Dim fbExisteDistribucionCIMG As Boolean
Dim lbITFCtaExonerada As Boolean 'FRHU 20150713 OBSERVACION
Dim bInstFinanc As Boolean 'FRHU 20150713 OBSERVACION
Dim sPersCod As String
Dim fnGastoAdminAdicional As Currency
'***************************************************************************
Private Sub Form_Load()
    fsVarPersCodCMAC = ""
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
    Call CargaParametros
    Limpiar
    CentraForm Me
End Sub
Private Sub cmdBuscar_Click()
    Dim loPers As COMDPersona.UCOMPersona
    Dim lsPersCod As String, lsPersNombre As String
    Dim lsEstados As String
    Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
    Dim lrCreditos As ADODB.Recordset
    Dim loCuentas As COMDPersona.UCOMProdPersona
    
    On Error GoTo ControlError
    
    Set loPers = New COMDPersona.UCOMPersona
        Set loPers = frmBuscaPersona.inicio
        If Not loPers Is Nothing Then
            lsPersCod = loPers.sPersCod
            sPersCod = loPers.sPersCod
            lsPersNombre = loPers.sPersNombre
        Else
            Exit Sub
        End If
    Set loPers = Nothing
    
    ' Selecciona Estados
    lsEstados = gColocEstTransferido
    
    If Trim(lsPersCod) <> "" Then
        Set loPersCredito = New COMDColocRec.DCOMColRecCredito
            Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
        Set loPersCredito = Nothing
    End If
    
    Set loCuentas = New COMDPersona.UCOMProdPersona
        Set loCuentas = frmProdPersona.inicio(lsPersNombre, lrCreditos)
        If loCuentas.sCtaCod <> "" Then
            AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
            AXCodCta.SetFocusCuenta
        End If
    Set loCuentas = Nothing
    'ventana = 1
    Exit Sub
    
ControlError:       ' Rutina de control de errores.
        MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
            " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaCredito (AXCodCta.NroCuenta)
End Sub
Private Sub cboModalidad_Click()
    Dim lnMontoChq As Double
    On Error GoTo ErrCboModalidad
    If Len(AXCodCta.NroCuenta) <> 18 Then Exit Sub
    txtnumdoc.Text = ""
    Set oDocRec = New UDocRec
    If Me.cboModalidad.ListIndex <> -1 Then
        If CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoCheque Then
            'By Capi 14042008 para que jale solo cheques valorizados caja general
            'MatDatos = frmBuscaCheque.BuscaCheque(gChqEstEnValorizacion, CInt(Mid(AXCodCta.NroCuenta, 9, 1)))
            'MatDatos = frmBuscaCheque.BuscaCheque(gChqEstValorizado, CInt(Mid(AXCodCta.NroCuenta, 9, 1)), 1)
            '
            '************************RECO 2013-07-23***************
            Dim lsVarCondicion As Boolean
            Dim oform As New frmChequeBusqueda 'EJVG20140228
            lsVarCondicion = False
            Do While lsVarCondicion = False
                'MatDatos = frmBuscaCheque.BuscaCheque(gChqEstValorizado, CInt(Mid(AXCodCta.NroCuenta, 9, 1)), 1)
                Set oDocRec = oform.iniciarBusqueda(Val(Mid(AXCodCta.NroCuenta, 9, 1)), TipoOperacionCheque.CRED_Pago, AXCodCta.NroCuenta)
                'If MatDatos(0) = "" Then
                If oDocRec.fsNroDoc = "" Then
                    Exit Do
                End If
                'If frmBuscaCheque.pnMontoDisponible < TxtTotalAPagar.Text Then
                'FRHU 20150715 OBSERVACION
                'lnMontoChq = DeducirMontoxITF(oDocRec.fnMonto)
                If lbITFCtaExonerada Or bInstFinanc Then
                    lnMontoChq = oDocRec.fnMonto
                Else
                    lnMontoChq = DeducirMontoxITF(oDocRec.fnMonto)
                End If
                'FIN FRHU 20150715
                If oDocRec.fnMonto < CCur(TxtTotalAPagar.Text) Then 'EJVG20140408
                    MsgBox "No es posible realizar el pago con ese cheque porque no cuenta con saldo suficiente para realizar la operación", _
                    vbInformation, " Aviso "
                'Exit Sub
                    
                Else
                    lsVarCondicion = True
                End If
            Loop
            Set oform = Nothing
            '***********************END RECO**********************

            'If MatDatos(0) <> "" Then
            If oDocRec.fsNroDoc <> "" Then 'EJVG20140228
                'txtNumDoc.Text = MatDatos(4)
                txtnumdoc.Text = oDocRec.fsNroDoc
                'Modificado por DAOR 20070809
                'AXMontoPago.Text = MatDatos(3)
                If Not fbExisteDistribucionCIMG Then
                    'By Capi 15042008
                    'AXMontoPago.Text = MatDatos(3)
                    'AXMontoPago.Text = MatDatos(0)
                    AXMontoPago.Text = lnMontoChq
                End If
            Else
                txtnumdoc.Text = ""
            End If
            txtnumdoc.Visible = True
        
        'RIRO20140530 ERS017 ***
        ElseIf CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoVoucher Then
        
            Dim oformV As frmCapRegVouDepBus
            Dim lnTipMot As Integer
            Dim sGlosa As String
            Dim sIF As String
            Dim sVaucher As String
            Dim sPersCod As String
            Dim sNombre As String
            Dim sDireccion As String
            Dim sDocumento As String
            Dim bCondicion As Boolean
                        
            Set oformV = New frmCapRegVouDepBus
            lnTipMot = 10 ' Pago Credito
                
            oformV.iniciarFormularioDeposito CInt(Mid(AXCodCta.NroCuenta, 9, 1)), lnTipMot, sGlosa, sIF, sVaucher, nMontoVoucher, sPersCod, nMovNroRVD, nMovNroRVDPen, sNombre, sDireccion, sDocumento, AXCodCta.NroCuenta
            If nMontoVoucher < AXMontoPago.value + CDbl(lblITF.Caption) Then
                If nMovNroRVD > 0 Then
                    MsgBox "No es posible realizar el pago con el Voucher porque no cuenta con saldo suficiente para realizar la operación", _
                    vbExclamation, "Aviso"
                End If
                nMovNroRVD = 0
                nMovNroRVDPen = -1
                nMontoVoucher = 0
                sNombre = ""
                sDireccion = ""
                If cboModalidad.Enabled Then cboModalidad.SetFocus
                Exit Sub
            Else
                If Len(sVaucher) = 0 Then
                    txtnumdoc.Text = sVaucher
                Else
                    txtnumdoc.Text = Trim(Mid(sVaucher, 1, Len(sVaucher) - 10))
                End If
                txtnumdoc.Visible = True
            End If
        'END RIRO *****
        Else
            txtnumdoc.Visible = False
            'MatDatos(0) = ""
        End If
    End If
    Exit Sub
ErrCboModalidad:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub cboModalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'If AXMontoPago.Enabled = True Then '**Juez 20120516 Agegado x observaciones en Pruebas 'FRHU 20150611
        If txtGlosa.Enabled = True Then '**Juez 20120516 Agegado x observaciones en Pruebas
            'AXMontoPago.SetFocus
            txtGlosa.SetFocus 'FRHU 20150611
        End If
    End If
End Sub
Private Sub BuscaCredito(ByVal psCodCta As String)
    Dim loValCred As COMDColocRec.DCOMColRecCredito
    Dim lrDatCredito As New ADODB.Recordset
    Dim lrCIMG As ADODB.Recordset 'DAOR 20070416
    'FRHU 20150713 OBSERVACION
    Dim loGrabar As COMNColocRec.NCOMColRecCredito
    Dim lsCtaAhorro As String
    Dim lsCodPersTitCtaAhorro As String
    Dim rs As ADODB.Recordset
    'FIN FRHU 20150713 OBSERVACION
    On Error GoTo ControlError
    
    'FRHU 20150713 OBSERVACION
    Set loGrabar = New COMNColocRec.NCOMColRecCredito
        Set rs = loGrabar.ObtenerCtaAhorroPorCredito(AXCodCta.NroCuenta)
        If Not rs.EOF And Not rs.BOF Then
            lsCtaAhorro = rs!CtaAhorro
            lsCodPersTitCtaAhorro = rs!cPersCod
        Else
            MsgBox "Cuenta de Ahorro de recaudo de la FOCMAC no existe", vbInformation, "Aviso"
            Exit Sub
        End If
        Set rs = Nothing
    Set loGrabar = Nothing
    lbITFCtaExonerada = fgITFVerificaExoneracion(lsCtaAhorro)
    Dim oDInstFinan As COMDPersona.DCOMInstFinac
    Set oDInstFinan = New COMDPersona.DCOMInstFinac
        bInstFinanc = oDInstFinan.VerificaEsInstFinanc(lsCodPersTitCtaAhorro)
    Set oDInstFinan = Nothing
    'FIN FRHU 20150713
    fnRegCancelacion = 0
    Set loValCred = New COMDColocRec.DCOMColRecCredito
        Set lrCIMG = loValCred.CredTransfObtieneDistribucionCIMGCobranza(psCodCta, gdFecSis)
        Set lrDatCredito = loValCred.ObtieneDatosPersCredTranferidos(psCodCta)
    Set loValCred = Nothing
    
    If Not lrDatCredito.BOF And Not lrDatCredito.EOF Then
        lblCliente.Caption = lrDatCredito!cPersNombre
        lblDOI.Caption = lrDatCredito!DOI
        LblMoneda.Caption = lrDatCredito!Moneda
        lblMetLiquid.Caption = lrDatCredito!MetLiquidacion
        fnSaldoCap = lrDatCredito!nSaldo
        fnSaldoIntComp = lrDatCredito!nIntComp
        fnSaldoIntMorat = lrDatCredito!nIntMora
        fnSaldoGasto = lrDatCredito!nGasto
        fnNroCalend = lrDatCredito!nNroCalen
        fnEstadoNew = lrDatCredito!nPrdEstado
        fnEstadoIni = lrDatCredito!nPrdEstado
        fnNroUltGastoCta = lrDatCredito!nUltGasto
        fnEstadoTransf = lrDatCredito!nPrdEstadoTransf
    Else
        MsgBox "No existe datos para este credito", vbInformation, "Aviso"
        Exit Sub
    End If
    lblCapital.Caption = Format(fnSaldoCap, "#,##0.00")
    lblInteres.Caption = Format(fnSaldoIntComp, "#,##0.00")
    lblMora.Caption = Format(fnSaldoIntMorat, "#,##0.00")
    lblGasto.Caption = Format(fnSaldoGasto, "#,##0.00")
    fnTotalDeuda = Format(fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat + fnSaldoGasto, "##0.00")
    
    If Not lrCIMG.EOF Then
        fnCapDist = lrCIMG!nCapital
        fnIntCompDist = lrCIMG!nIntComp
        fnIntMoratDist = lrCIMG!nMora
        fnGastoDist = lrCIMG!nGasto
        fbExisteDistribucionCIMG = True
        AXMontoPago.Text = Format(fnCapDist + fnIntCompDist + fnIntMoratDist + fnGastoDist, "#0.00")
        AXMontoPago.Enabled = False
        If Not IsNull(lrCIMG!nRegCancelacion) Then
            fnRegCancelacion = lrCIMG!nRegCancelacion
            'MsgBox "Al grabar el pago, el crédito será cancelado como lo estableció en el Area de Recuperaciones", vbInformation, "Aviso"
            MsgBox "Al grabar el pago, el crédito será cancelado como lo estableció el Area de Recuperaciones", vbInformation, "Aviso" 'FRHU 20150612
        End If
        Call AXMontoPago_KeyPress(13)
    Else
        fbExisteDistribucionCIMG = False
    End If
Exit Sub
ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
           "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
Private Sub cmdGrabar_Click()
    Dim loContFunct As COMNContabilidad.NCOMContFunciones 'NContFunciones
    Dim loGrabar As COMNColocRec.NCOMColRecCredito 'NColRecCredito
    Dim loImprime As COMNColocRec.NCOMColRecImpre 'NColRecImpre
    Dim loPrevio As previo.clsprevio
    Dim lnDocTpo As Integer
    Dim lsNroDoc As String
    Dim lsMovNro As String
    Dim loMov As COMDMov.DCOMMov
    Dim lsOpeCod As String
    Dim sOpeCod As String
    Dim lsFechaHoraGrab As String
    Dim lsCadImprimir As String
    Dim lsNombreCliente As String
    Dim lsCtaAhorro As String
    Dim rs As ADODB.Recordset
    Dim pMatDatosAho As Variant
On Error GoTo ControlError
    
    If Len(Trim(lblMetLiquid)) <> 4 Then
        MsgBox "Metodo de Liquidación no válido o no definido", vbInformation, "Aviso"
        Exit Sub
    End If
    lsNombreCliente = Mid(Me.lblCliente.Caption, 1, 30)
    If cboModalidad = "" Then
        MsgBox "Seleccione modalidad de Pago", vbInformation, "Aviso"
        Me.cboModalidad.SetFocus
        Exit Sub
    End If
    'FRHU 20150612 'Observacion
    If txtGlosa.Text = "" Then
        MsgBox "Por favor, ingrese una glosa.", vbInformation, "Aviso"
        Me.txtGlosa.SetFocus
        Exit Sub
    End If
    'FIN FRHU 20150612
    If CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoCheque Then
        lsOpeCod = gColRecOpePagTransfFocMacmChq
        If txtnumdoc = "" Then
            MsgBox "Numero de Cheque no Válido", vbInformation, "Aviso"
            Exit Sub
        End If
        lnDocTpo = TpoDocCheque
        lsNroDoc = Trim(txtnumdoc)
    ElseIf CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoVoucher Then
        lnDocTpo = 0
        lsNroDoc = Trim(txtnumdoc)
        lsOpeCod = gColRecOpePagTransfFocMacmVou
        If nMovNroRVD = 0 Then
            MsgBox "Debe seleccionar un voucher para continuar con la operacion", vbInformation, "Aviso"
            Exit Sub
        End If
        If nMontoVoucher < AXMontoPago.value + CDbl(lblITF.Caption) Then
            If nMovNroRVD > 0 Then
                MsgBox "No es posible realizar el pago con el Voucher porque no cuenta con saldo suficiente para realizar la operación", _
                vbExclamation, "Aviso"
            End If
            If AXMontoPago.Enabled Then AXMontoPago.SetFocus
            Exit Sub
        End If
    Else
        lnDocTpo = 0
        lsNroDoc = ""
        lsOpeCod = gColRecOpePagTransfFocMacmEfe
    End If
    If CCur(AXMontoPago.Text) > CCur(fnTotalDeuda) Then
        MsgBox "Monto a Pagar no debe Exceder el Total de Deuda", vbInformation, "Aviso"
        Me.AXMontoPago.SetFocus
        Exit Sub
    End If
    If CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoCheque Then
        If Not ValidaSeleccionCheque Then Exit Sub
        If CCur(TxtTotalAPagar.Text) > oDocRec.fnMonto Then
            MsgBox "Disponible del cheque no cubre el Monto a Pagar", vbInformation, "Aviso"
            If AXMontoPago.Visible And AXMontoPago.Enabled Then Me.AXMontoPago.SetFocus
            Exit Sub
        End If
    End If
    Set loGrabar = New COMNColocRec.NCOMColRecCredito
        Set rs = loGrabar.ObtenerCtaAhorroPorCredito(AXCodCta.NroCuenta)
        If Not rs.EOF And Not rs.BOF Then
            lsCtaAhorro = rs!CtaAhorro
        Else
            'MsgBox "Cuenta de Ahorro de la FOCMACM no existe", vbInformation, "Aviso"
            MsgBox "Cuenta de Ahorro de recaudo de la FOCMAC no existe", vbInformation, "Aviso" 'FRHU 20150713 OBSERVACION
            Exit Sub
        End If
        Set rs = Nothing
    Set loGrabar = Nothing
    Call CalculaDistribucionPago
    If fnNewSaldoCap <= 10 And fbExisteDistribucionCIMG = False Then
        MsgBox "Para realizar este pago se requiere Autorización de Recuperaciones", vbInformation, "Aviso"
        Exit Sub
    End If
    'inicializa Datos de Ahorros
    ReDim pMatDatosAho(17)
    pMatDatosAho(0) = "" 'Cuenta de Ahorros
    pMatDatosAho(1) = "0.00" 'Monto de Apertura
    pMatDatosAho(2) = "0.00" 'Interes Ganado de Abono
    pMatDatosAho(3) = "0.00" 'Interes Ganado de Retiro Gastos
    pMatDatosAho(4) = "0.00" 'Interes Ganado de Retiro Cancelaciones
    pMatDatosAho(5) = "0.00" 'Monto de Abono
    pMatDatosAho(6) = "0.00" 'Monto de Retiro de Gastos
    pMatDatosAho(7) = "0.00" 'Monto de Retiro de Cancelaciones
    pMatDatosAho(8) = "0.00" 'Saldo Disponible Abono
    pMatDatosAho(9) = "0.00" 'Saldo Contable Abono
    pMatDatosAho(10) = "0.00" 'Saldo Disponible Retiro de Gastos
    pMatDatosAho(11) = "0.00" 'Saldo Contable Retiro de Gastos
    pMatDatosAho(12) = "0.00" 'Saldo Disponible Retiro de Cancelaciones
    pMatDatosAho(13) = "0.00" 'Saldo Contable Retiro de Cancelaciones
    pMatDatosAho(14) = "0.00" 'Monto de Retiro Efectivo 'FRHU20140228 RQ14006
    pMatDatosAho(15) = "0.00" 'Saldo Disponible Retiro Efectivo 'FRHU20140228 RQ14006
    pMatDatosAho(16) = "0.00" 'Saldo Contable Retiro Efectivo 'FRHU20140228 RQ14006
    pMatDatosAho(17) = "0.00" 'Interes Ganado Retiro Efectivo 'FRHU20140228 RQ14006
    'If MsgBox(" Desea Grabar Pago de Credito Transferido FOCMACM ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    If MsgBox(" ¿Desea grabar el pago del crédito transferido FOCMAC? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then 'FRHU 20150612 Observacion
        If oDocRec Is Nothing Then Set oDocRec = New UDocRec 'EJVG20140408
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabar = New COMNColocRec.NCOMColRecCredito
            'RIRO20140620 ERS017 *********
            If CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoCheque Then
                sOpeCod = "990107"
            ElseIf CInt(Trim(Right(cboModalidad.Text, 10))) = gColocTipoPagoVoucher Then
                sOpeCod = "990121"
            Else
                sOpeCod = "990105"
            End If
            Call loGrabar.nPagoCreditoTransferido(AXCodCta.NroCuenta, lsFechaHoraGrab, lsOpeCod, _
                 lsMovNro, CCur(Me.AXMontoPago.Text), Me.lblMetLiquid.Caption, fnNewSaldoCap, fnNewSaldoIntComp, _
                 fnNewSaldoIntMorat, fnNewSaldoGasto, fnNroUltGastoCta, _
                 fnCapPag, fnIntCompPag, fnIntMoratPag, fnGastoPag, fmMatGastos, _
                 fnNroCalend, fnEstadoNew, False, , CDbl(lblITF.Caption), oDocRec.fnTpoDoc, oDocRec.fsNroDoc, sOpeCod, fsVarPersCodCMAC, , , , , , , gnMovNro, oDocRec.fsPersCod, oDocRec.fsIFTpo, oDocRec.fsIFCta, nMovNroRVD, nMovNroRVDPen, lsCtaAhorro, pMatDatosAho, fnEstadoTransf) 'EJVG20140408
            'END RIRO *********************
            If gnMovNro = 0 Then
                MsgBox "La operación no se realizó, favor intente nuevamente", vbInformation, "Aviso"
                Exit Sub
            End If
            loGrabar.nActualizarColocTransfAutorizado (AXCodCta.NroCuenta)
        Set loGrabar = Nothing
        
        Set loMov = New COMDMov.DCOMMov
            If gbITFAplica = True And CCur(lblITF.Caption) > 0 Then
               Call loMov.InsertaMovRedondeoITF(lsMovNro, 1, CCur(lblITF.Caption) + nRedondeoITF, CCur(lblITF.Caption)) 'BRGO 20110914
            End If
        Set loMov = Nothing
        
        'Impresión
        Set loImprime = New COMNColocRec.NCOMColRecImpre
            'FRHU 20150603 ERS022-2015: Se agrego "lsOpeCod" en la función
            lsCadImprimir = loImprime.nPrintReciboPagoCredRecup(gsNomAge, lsFechaHoraGrab, AXCodCta.NroCuenta, _
            lsNombreCliente, CCur(Me.AXMontoPago.Text), gsCodUser, " ", CDbl(lblITF.Caption), gImpresora, gbImpTMU, lsOpeCod)
        Set loImprime = Nothing
        
        Set loPrevio = New previo.clsprevio
            loPrevio.PrintSpool sLpt, lsCadImprimir, True, 22
            Do While True
                'If MsgBox("Reimprimir Recibo de Pago de Credito en Recuperaciones ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                If MsgBox("¿Desea reimprimir el recibo de pago del crédito? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then 'FRHU 20150612 Observacion
                    loPrevio.PrintSpool sLpt, lsCadImprimir, True, 22
                Else
                    Exit Do
                End If
            Loop
        Set loPrevio = Nothing
        
        Limpiar
        AXCodCta.Enabled = True
        AXCodCta.SetFocus
        'INICIO JHCU ENCUESTA 16-10-2019
        Encuestas gsCodUser, gsCodAge, "ERS0292019", sOpeCod
        'FIN
    Else
        MsgBox " Grabación cancelada ", vbInformation, " Aviso "
    End If
    
    Exit Sub
ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
Private Sub CalculaDistribucionPago()
    fnNewSaldoCap = 0: fnNewSaldoIntComp = 0: fnNewSaldoIntMorat = 0: fnNewSaldoGasto = 0
    fnCapPag = 0: fnIntCompPag = 0: fnIntMoratPag = 0: fnGastoPag = 0
    ' Distribuye Monto
    Call DistribuyePago
    
    fnNewSaldoCap = fnSaldoCap - fnCapPag
    fnNewSaldoIntComp = fnSaldoIntComp - fnIntCompPag
    fnNewSaldoIntMorat = fnSaldoIntMorat - fnIntMoratPag
    fnNewSaldoGasto = fnSaldoGasto - fnGastoPag

    'fnNewSaldoIntCompGen = fnIntCompGenerado
    'fnNewSaldoIntMoraGen = fnIntMoraGenerado
   
    If fnRegCancelacion = 1 Then
        fnEstadoNew = 2305
    Else
        fnEstadoNew = fnEstadoIni
    End If
End Sub
Private Sub DistribuyePago()
    Dim lnMontoDistrib As Currency
    
    fnMontoPagar = Format(Me.AXMontoPago.Text, "#0.00")
    
    If fbExisteDistribucionCIMG Then
        Call EstablecerCIMGPersonalizado
    Else
        Call ReCalCulaGasto(AXCodCta.NroCuenta)
        Call DistribuyePagosSegunMetLiquidacion 'DistribuyePagoCominAbogado
    End If
    fnGastoAdminAdicional = 0
    lnMontoDistrib = Format(Me.AXMontoPago.Text, "#0.00")
    
    If lnMontoDistrib >= (fnSaldoCap + fnSaldoIntComp + fnSaldoGasto + fnSaldoIntMorat) Then
        If fnMontoPagar > 0 Then
            fnGastoAdminAdicional = Format(fnMontoPagar, "#0.00")
        End If
    End If
End Sub
'** Procedimiento que establece los montos CIMG distribuidos de forma manual
Private Sub EstablecerCIMGPersonalizado()
    Dim lnGastoDistrib As Currency
    Dim i As Integer
    fnCapPag = fnCapDist
    fnIntCompPag = fnIntCompDist
    fnIntMoratPag = fnIntMoratDist
    fnGastoPag = fnGastoDist
    fnMontoPagar = 0
    lnGastoDistrib = fnGastoPag
    
    Call ReCalCulaGasto(AXCodCta.NroCuenta)
    
    For i = 0 To UBound(fmMatGastos) - 1
        If CInt(fmMatGastos(i, 4)) = gColRecGastoEstPendiente And lnGastoDistrib > 0 _
           And (CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))) > 0 Then
            If lnGastoDistrib >= (CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))) Then
                lnGastoDistrib = lnGastoDistrib - (CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3)))
                'Actualiza el monto Pagado
                'fmMatGastos(i, 3) = CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))
                'Actualiza el estado del gasto
                fmMatGastos(i, 4) = gColRecGastoEstPagado
                fmMatGastos(i, 5) = "S" ' Si se ha modificado
                fmMatGastos(i, 6) = CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))
            Else
                'Actualiza el monto pagado
                fmMatGastos(i, 3) = Format(CDbl(fmMatGastos(i, 3)) + lnGastoDistrib, "#0.00")
                fmMatGastos(i, 4) = gColRecGastoEstPendiente
                fmMatGastos(i, 5) = "S" ' Si se ha modificado
                fmMatGastos(i, 6) = Format(lnGastoDistrib, "#0.00")
                lnGastoDistrib = 0
            End If
        End If
    Next i
    
End Sub
Private Sub ReCalCulaGasto(ByVal psCtaCod As String)
    Dim lrDatGastos As New ADODB.Recordset
    Dim loValCred As COMDColocRec.DCOMColRecCredito
    Dim lsmensaje As String
    
    Set loValCred = New COMDColocRec.DCOMColRecCredito
    Set lrDatGastos = loValCred.dObtieneListaGastosxCredito(psCtaCod, lsmensaje, True)
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    If lrDatGastos Is Nothing Then   ' Hubo un Error
        Exit Sub
    End If
    
    Set fmMatGastos = Nothing
    
        Dim i As Integer
        ReDim fmMatGastos(0)
        ReDim fmMatGastos(lrDatGastos.RecordCount, 11)
        
        Do While Not lrDatGastos.EOF
            If lrDatGastos!nColocRecGastoEstado = gColRecGastoEstPendiente Then
                fmMatGastos(i, 1) = lrDatGastos!nNroGastoCta
                fmMatGastos(i, 2) = lrDatGastos!nMonto
                fmMatGastos(i, 3) = lrDatGastos!nMontoPagado
                fmMatGastos(i, 4) = lrDatGastos!nColocRecGastoEstado
                fmMatGastos(i, 5) = "N" ' Estado del Gasto
                fmMatGastos(i, 6) = 0 '(fmMatGastos(i, 2) - fmMatGastos(i, 3)) 'avmm 0   ' Monto a Cubrir del Gasto
                fmMatGastos(i, 7) = lrDatGastos!nPrdConceptoCod
                i = i + 1
            End If
            lrDatGastos.MoveNext
        Loop
    
End Sub
Private Sub DistribuyePagosSegunMetLiquidacion()
    Dim lsPrio1 As String, lsPrio2 As String, lsPrio3 As String, lsPrio4 As String
    
    lsPrio1 = Mid(Me.lblMetLiquid, 1, 1)
    lsPrio2 = Mid(Me.lblMetLiquid, 2, 1)
    lsPrio3 = Mid(Me.lblMetLiquid, 3, 1)
    lsPrio4 = Mid(Me.lblMetLiquid, 4, 1)
    
    fnCapPag = 0: fnIntCompPag = 0: fnIntMoratPag = 0:  fnGastoPag = 0
    
    If fnMontoPagar > 0 Then
        Select Case lsPrio1
            Case "G": Call CubrirGastos
            Case "M": Call CubrirMora
            Case "I": Call CubrirInteres
            Case "C": Call CubrirCapital
        End Select
    End If
    
    If fnMontoPagar > 0 Then
        Select Case lsPrio2
            Case "G": Call CubrirGastos
            Case "M": Call CubrirMora
            Case "I": Call CubrirInteres
            Case "C": Call CubrirCapital
        End Select
    End If
    
    If fnMontoPagar > 0 Then
        Select Case lsPrio3
            Case "G":  Call CubrirGastos
            Case "M":  Call CubrirMora
            Case "I":  Call CubrirInteres
            Case "C":  Call CubrirCapital
        End Select
    End If
    
    If fnMontoPagar > 0 Then
        Select Case lsPrio4
            Case "G": Call CubrirGastos
            Case "M": Call CubrirMora
            Case "I": Call CubrirInteres
            Case "C": Call CubrirCapital
        End Select
    End If
End Sub
Private Sub CubrirCapital()
    'Cubro Capital
    If fnSaldoCap > 0 And fnMontoPagar > 0 Then
        If fnMontoPagar >= fnSaldoCap Then
            fnCapPag = fnSaldoCap
            fnMontoPagar = fnMontoPagar - fnSaldoCap
        Else
            fnCapPag = fnMontoPagar
            fnMontoPagar = 0
        End If
    End If
End Sub
Private Sub CubrirInteres()
    'Cubro Interes
    If fnSaldoIntComp > 0 And fnMontoPagar > 0 Then
        If fnMontoPagar >= fnSaldoIntComp Then
            fnIntCompPag = fnSaldoIntComp
            fnMontoPagar = fnMontoPagar - fnSaldoIntComp
        Else
            fnIntCompPag = fnMontoPagar
            fnMontoPagar = 0
        End If
    End If
End Sub
Private Sub CubrirGastos()
    Dim lnGastoDistrib As Currency
    Dim i As Integer
    'Cubro Gastos
    If fnSaldoGasto > 0 And fnMontoPagar > 0 Then
        If fnMontoPagar >= fnSaldoGasto Then
            fnGastoPag = fnSaldoGasto
            fnMontoPagar = fnMontoPagar - fnSaldoGasto
        Else
            fnGastoPag = fnMontoPagar
            fnMontoPagar = 0
        End If
    End If
    'Actualiza la Matriz de Gastos
    lnGastoDistrib = fnGastoPag
    
    For i = 0 To UBound(fmMatGastos) - 1
        If CInt(fmMatGastos(i, 4)) = gColRecGastoEstPendiente And lnGastoDistrib > 0 _
           And (CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))) > 0 Then
            If lnGastoDistrib >= (CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))) Then
                lnGastoDistrib = lnGastoDistrib - (CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3)))
                fmMatGastos(i, 4) = gColRecGastoEstPagado
                fmMatGastos(i, 5) = "S" ' Si se ha modificado
                fmMatGastos(i, 6) = CDbl(fmMatGastos(i, 2)) - CDbl(fmMatGastos(i, 3))
            Else
                'Actualiza el monto pagado
                fmMatGastos(i, 3) = Format(CDbl(fmMatGastos(i, 3)) + lnGastoDistrib, "#0.00")
                fmMatGastos(i, 4) = gColRecGastoEstPendiente
                fmMatGastos(i, 5) = "S" ' Si se ha modificado
                fmMatGastos(i, 6) = Format(lnGastoDistrib, "#0.00")
                lnGastoDistrib = 0
            End If
        End If
    Next i
    
End Sub
Private Sub CubrirMora()
    'Cubro Mora
    If fnSaldoIntMorat > 0 And fnMontoPagar > 0 Then
        If fnMontoPagar >= fnSaldoIntMorat Then
            fnIntMoratPag = fnSaldoIntMorat
            fnMontoPagar = fnMontoPagar - fnSaldoIntMorat
        Else
            fnIntMoratPag = fnMontoPagar
            fnMontoPagar = 0
        End If
    End If
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    Limpiar
    AXCodCta.SetFocusAge
End Sub
Private Sub AXMontoPago_KeyPress(KeyAscii As Integer)
    Dim lnITF As Double
    If KeyAscii = 13 Then
    'FRHU 20150713 OBSERVACION
'        If val(AXMontoPago.Text) = 0 Then
'            Exit Sub
'        End If
'        lnITF = gITF.fgITFCalculaImpuesto(CCur(AXMontoPago.Text))
'        nRedondeoITF = fgDiferenciaRedondeoITF(lnITF)
'        If nRedondeoITF > 0 Then
'            lnITF = lnITF - nRedondeoITF
'        End If
'        LblITF = Format(lnITF, "#,##0.00")
'        TxtTotalAPagar = Format(CCur(AXMontoPago.Text) + lnITF, "#0.00")
    'FIN FRHU 20150713
        'CalculaDistribucionPago
        cmdGrabar.Enabled = True
        If cmdGrabar.Enabled And cmdGrabar.Visible Then cmdGrabar.SetFocus
    End If
End Sub
'FRHU 20150713 OBSERVACION
Private Sub AXMontoPago_Change()
Dim lnITF As Double
    If Val(AXMontoPago.Text) = 0 Then
        Exit Sub
    End If
    If Not lbITFCtaExonerada Then
        lnITF = gITF.fgITFCalculaImpuesto(CCur(AXMontoPago.Text))
        nRedondeoITF = fgDiferenciaRedondeoITF(lnITF)
        If nRedondeoITF > 0 Then
            lnITF = lnITF - nRedondeoITF
        End If
        lblITF = Format(lnITF, "#,##0.00")
        If bInstFinanc Then
            lblITF.Caption = "0.00"
            lnITF = 0
        End If
    Else
        lnITF = 0
        lblITF.Caption = "0.00"
    End If
    TxtTotalAPagar = Format(CCur(AXMontoPago.Text) + lnITF, "#0.00")
End Sub
'FIN FRHU 20150713
Private Sub CargaParametros()
    Dim oConstante As New COMDConstSistema.DCOMGeneral
    Dim rsConstante As New ADODB.Recordset
    
    Set rsConstante = oConstante.GetConstante(gColocTipoPago, , "'[125]'")
    CargaCombo cboModalidad, rsConstante
    Set rsConstante = Nothing
    Set oConstante = Nothing
End Sub
Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
    lblCliente.Caption = ""
    lblDOI.Caption = ""
    LblMoneda.Caption = ""
    lblMetLiquid.Caption = ""
    lblCapital.Caption = ""
    lblInteres.Caption = ""
    lblMora.Caption = ""
    lblGasto.Caption = ""
    
    cboModalidad.ListIndex = -1
    txtnumdoc.Text = ""
    txtGlosa.Text = ""
    Me.AXMontoPago.Text = 0
    Me.TxtTotalAPagar = 0
    Me.lblITF = 0
    cmdGrabar.Enabled = False 'FRHU 20150611 Observacion
    'FRHU 20150713 OBSERVACION
    lbITFCtaExonerada = False
    bInstFinanc = False
    'FIN FRHU 20150713
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If AXMontoPago.Enabled = True Then
            AXMontoPago.SetFocus
        End If
    End If
End Sub
Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtGlosa.SetFocus
End Sub
Private Sub TxtTotalAPagar_GotFocus()
    fEnfoque TxtTotalAPagar
End Sub
Private Sub TxtTotalAPagar_KeyPress(KeyAscii As Integer)
    Call AXMontoPago_KeyPress(KeyAscii)
End Sub
Private Function ValidaSeleccionCheque() As Boolean
    ValidaSeleccionCheque = True
    If oDocRec Is Nothing Then
        ValidaSeleccionCheque = False
        Exit Function
    End If
    If Len(Trim(oDocRec.fsNroDoc)) = 0 Then
        ValidaSeleccionCheque = False
        Exit Function
    End If
End Function
Private Sub Form_Unload(Cancel As Integer)
    Set oDocRec = Nothing
End Sub
