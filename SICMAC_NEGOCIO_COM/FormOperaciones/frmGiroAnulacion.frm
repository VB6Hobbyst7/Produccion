VERSION 5.00
Begin VB.Form frmGiroAnulacion 
   Caption         =   "Giros - Anulación de Giro"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   Icon            =   "frmGiroAnulacion.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7920
      TabIndex        =   12
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmbExaminar 
      Caption         =   "Examinar"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Giro"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   9135
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4440
         TabIndex        =   15
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Apertura:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Giro:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Ag. Destino:"
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblFecAper 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblTpoGiro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblAgenciaDes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4440
         TabIndex        =   3
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lblMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4920
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
   End
   Begin SICMACT.ActXCodCta txtCuenta 
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   767
      Texto           =   "Giro N°"
      EnabledCta      =   -1  'True
      EnabledAge      =   -1  'True
      Prod            =   "239"
      CMAC            =   "109"
   End
   Begin VB.Label lblMoneda 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   960
      TabIndex        =   16
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "Comisión:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1830
      Width           =   855
   End
   Begin VB.Label lblMontoCom 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1560
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "frmGiroAnulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmGiroAnulacion
'** Descripción : Formulario para realizar la anulación de giros
'** Creación : RECO, 20140410 - ERS008-2014
'**********************************************************************************************

Option Explicit
Dim nmoneda As Integer
Dim sOperacion As String, sRemitente As String
Dim sDestinatarioNomb As String
Dim GClaveGiro As String
Public lnValPinPad As Integer

Private Sub cmdAceptar_Click()
    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII
    
    Dim clsGiro As COMNCaptaServicios.NCOMCaptaServicios
    Set clsGiro = New COMNCaptaServicios.NCOMCaptaServicios
    Dim lsBoleta As String, nMontoComis As String
    
    Dim lnMovNroRef As Long
    Dim lnMovNro As Long
    
    Dim oMov As COMDMov.DCOMMov
    Set oMov = New COMDMov.DCOMMov
    
    nMontoComis = Val(lblMontoCom.Caption)
    
    If MsgBox("¿Desea Grabar la Operación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrGraba
    Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim sMovNro As String, sCuentaGiro As String, sPersLavDinero As String
    Dim ClsMov As COMNContabilidad.NCOMContFunciones
    
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion
        
    Set ClsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set ClsMov = Nothing
    
    
    lnMovNroRef = clsGiro.ServGiroAnulacion(Me.txtCuenta.NroCuenta, sMovNro, lblMonto.Caption, nMontoComis, sDestinatarioNomb, lsBoleta, sRemitente, Me.lblAgenciaDes.Caption)
    'Call clsGiro.ServGiroAnulacion(Me.txtCuenta.NroCuenta, sMovNro, lblMonto.Caption, nMontoComis, sDestinatarioNomb, lsBoleta, sRemitente, Me.lblAgenciaDes.Caption)
    
    'RECO20140724 ********************************************************
    Set ClsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set ClsMov = Nothing
    
    lnMovNro = clsGiro.ServGiroComision(Me.txtCuenta.NroCuenta, sMovNro, nMontoComis, "GIRO - Comisión Anulación Giro Cuenta : ", gServGiroComiAnul, gComServAnulGiro)
    
    Dim oBase As COMDCredito.DCOMCredActBD
    Set oBase = Nothing
    Set oBase = New COMDCredito.DCOMCredActBD
    Call oBase.dInsertMovRef(lnMovNroRef, lnMovNro)
    'RECO FIN*************************************************************
    If Trim(lsBoleta) <> "" Then
        Dim lbok As Boolean
        lbok = True
        Do While lbok
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
            Print #nFicSal, lsBoleta
            Print #nFicSal, ""
            Close #nFicSal
            If MsgBox("Desea Reimprimir Boleta ??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbok = False
            End If
        Loop
    End If
    Call LimpiarFormulario
    'INICIO JHCU ENCUESTA 16-10-2019
    Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
    'FIN
Exit Sub
ErrGraba:
    MsgBox err.Description, vbExclamation, "Error"
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCuenta As String, sMoneda As String
        sCuenta = txtCuenta.NroCuenta
        CargaDatosGiro sCuenta
    End If
End Sub

Private Sub cmbExaminar_Click()
    frmGiroPendiente.inicio frmGiroAnulacion
    Dim sCuenta As String
    Dim nlMoneda As Moneda
    
    sCuenta = txtCuenta.NroCuenta
    If Len(sCuenta) = 18 Then
        txtCuenta.SetFocusCuenta
        nlMoneda = CLng(Mid(sCuenta, 9, 1))
        nmoneda = nlMoneda
        If nlMoneda = COMDConstantes.gMonedaExtranjera Then
            'lblMonto.BackColor = &HC0FFC0
            LblMoneda.Caption = "US$"
        Else
            'lblMonto.BackColor = &HFFFFFF
            'lblMoneda.Caption = "S/"
            LblMoneda.Caption = gcPEN_SIMBOLO
            
        End If
        'SendKeys "{Enter}"
    End If
End Sub

Private Sub CargaDatosGiro(ByVal sCuenta As String)
    Dim rsGiro As ADODB.Recordset
    Dim rsGiroCom As ADODB.Recordset
    Dim clsGiro As COMNCaptaServicios.NCOMCaptaServicios
    Dim nFila As Long
    Dim sDestinatario As String
    Set clsGiro = New COMNCaptaServicios.NCOMCaptaServicios
    Set rsGiro = clsGiro.GetGiroDatos(sCuenta)
    
    Set rsGiroCom = New ADODB.Recordset
    
    If Not (rsGiro.EOF And rsGiro.BOF) Then
        sDestinatario = ""
        lblAgenciaDes = Trim(rsGiro("cAgencia"))
        lblMonto = Format$(rsGiro("nSaldo"), "#,##0.00")
        lblTpoGiro = Trim(rsGiro("cTipo"))
        lblFecAper = Format$(rsGiro("dPrdEstado"), "dd mmm yyyy")
        sRemitente = Trim(rsGiro("cRemitente"))
        sDestinatarioNomb = Trim(rsGiro("cDestinatario"))
        
        Dim dlsMant As COMDCaptaGenerales.DCOMCaptaGenerales
        Set dlsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    
        If dlsMant.GetNroOPeradoras(gsCodAge) > 1 Then
            If sRemitente = gsCodPersUser Then
                MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                Unload Me
                Exit Sub
            End If
        End If
        Set dlsMant = Nothing
        GClaveGiro = clsGiro.GetGiroSeguridad(sCuenta)
        If GClaveGiro <> "" Then
            'cmdClave.Enabled = True
        End If
        
        Dim nlMoneda As Moneda
    
        sCuenta = txtCuenta.NroCuenta
        If Len(sCuenta) = 18 Then
            txtCuenta.SetFocusCuenta
            nlMoneda = CLng(Mid(sCuenta, 9, 1))
            nmoneda = nlMoneda
            If nlMoneda = COMDConstantes.gMonedaExtranjera Then
                'lblMonto.BackColor = &HC0FFC0
                LblMoneda.Caption = "US$"
            Else
                'lblMonto.BackColor = &HFFFFFF
                'lblMoneda.Caption = "S/"
                LblMoneda.Caption = gcPEN_SIMBOLO
            End If
            'SendKeys "{Enter}"
        End If
    Else
        MsgBox "Número de Giro no encontrado o Cancelado.", vbInformation, "SICMACM - Aviso"
        txtCuenta.Age = ""
        txtCuenta.Cuenta = ""
        txtCuenta.SetFocusAge
        sRemitente = ""
    End If
    
    Set rsGiroCom = clsGiro.RecuperaValorComisionTarGiro(2)
    If Not (rsGiroCom.EOF And rsGiroCom.BOF) Then
        lblMontoCom = Format(IIf(nmoneda = 1, rsGiroCom!nMontoMN, rsGiroCom!nMontoME), "#,##0.00")
    Else
        MsgBox "No se encontró valor de comisión. Comuníquese con el departamento de TI", vbInformation, "SICMACM - Aviso"
        txtCuenta.Age = ""
        txtCuenta.Cuenta = ""
        txtCuenta.SetFocusAge
        sRemitente = ""
    End If
    Set clsGiro = Nothing
End Sub
Public Sub LimpiarFormulario()
    txtCuenta.NroCuenta = ""
    txtCuenta.CMAC = "109"
    txtCuenta.Prod = "239"
    lblFecAper.Caption = ""
    lblTpoGiro.Caption = ""
    lblAgenciaDes.Caption = ""
    lblMonto.Caption = ""
    LblMoneda.Caption = ""
    lblMontoCom.Caption = ""
End Sub
