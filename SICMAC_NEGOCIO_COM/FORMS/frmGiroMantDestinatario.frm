VERSION 5.00
Begin VB.Form frmGiroMantDestinatario 
   Caption         =   "Giros - Mantenimiento - Cambio de Destinatario"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   Icon            =   "frmGiroMantDestinatario.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   29
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7920
      TabIndex        =   28
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame frmDest 
      Enabled         =   0   'False
      Height          =   2175
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   6015
      Begin VB.TextBox txtReferencia 
         Height          =   285
         Left            =   1320
         TabIndex        =   26
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox txtDNI 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   960
         Width           =   4575
      End
      Begin SICMACT.TxtBuscar txtBuscaPersona 
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
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
         TipoBusPers     =   1
      End
      Begin VB.CheckBox chkClienteCMAC 
         Caption         =   "Cliente CMAC Maynas"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "  Destinatario"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Referencia"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "DNI:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame frmClave 
      Caption         =   "Cambio de Clave"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   6240
      TabIndex        =   13
      Top             =   2760
      Width           =   3015
      Begin VB.TextBox txtClaveNew 
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtClaveAnt 
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Nueva Clave:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Clave Giro:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame frmAgencia 
      Enabled         =   0   'False
      Height          =   855
      Left            =   6240
      TabIndex        =   11
      Top             =   1800
      Width           =   3015
      Begin VB.ComboBox cboAgencia 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "  Agencia"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Giro"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   9135
      Begin VB.Label lblMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4440
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblAgenciaDes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4440
         TabIndex        =   9
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lblTpoGiro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblFecAper 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Ag. Destino:"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Giro:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Apertura:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmbExaminar 
      Caption         =   "Examinar"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   1215
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
   Begin VB.Label lblMontoCom 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1560
      TabIndex        =   31
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblMoneda 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   960
      TabIndex        =   30
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "Comisión:"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   4130
      Width           =   855
   End
End
Attribute VB_Name = "frmGiroMantDestinatario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmGiroMantDestinatario
'** Descripción : Formulario para dar mantenimeinto al destinatario o agencia destino del Giro
'** Creación : RECO, 20140410 - ERS008-2014
'**********************************************************************************************

Option Explicit
Dim nmoneda As Integer
Dim sOperacion As String, sRemitente As String
Dim sCodRemitente As String
Dim SNroIdentidad As String
Dim GClaveGiro As String
Dim nTpoCliente As Integer
Dim sRemitenteNomb As String
Dim nTpoClienteAnt As Integer
Dim nTpoClienteAct As Integer
Dim sPersCodAnt As String
Dim sPersCodAct As String
Dim sNumDOIAnt As String
Dim sNumDOIAct As String
Dim fcAgeCodDest As String

Private Sub chkClienteCMAC_Click()
    If chkClienteCMAC.value = 0 Then
        Call HabilitarControles(True)
    Else
        Call HabilitarControles(False)
    End If
End Sub

Private Sub cmbExaminar_Click()
    frmGiroPendiente.Inicio frmGiroMantDestinatario
    Dim sCuenta As String
    Dim nlMoneda As Moneda
    
    sCuenta = txtCuenta.NroCuenta
    If Len(sCuenta) = 18 Then
        txtCuenta.SetFocusCuenta
        nlMoneda = CLng(Mid(sCuenta, 9, 1))
        nmoneda = nlMoneda
        If nlMoneda = COMDConstantes.gMonedaExtranjera Then
            'lblMonto.BackColor = &HC0FFC0
            lblMoneda.Caption = "US$"
        Else
            'lblMonto.BackColor = &HFFFFFF
            lblMoneda.Caption = "S/."
        End If
        'SendKeys "{Enter}"
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim fnCondicion As Integer
    Dim nMontoComis As Double
    Dim sCuenta As String
    Dim nmoneda As Moneda
    Dim nTipo As COMDConstantes.ProductoCuentaTipo
    Dim sAgenciaDest As String
    Dim rsDest As New ADODB.Recordset
    Dim lsClaveSeg As String
    Dim oConA As COMDConstSistema.DCOMUAcceso
    Dim nFicSal As String
    Dim lsBoleta As String
    Dim lsBoletaITF As String
    
    Dim objPersona As COMDPersona.DCOMPersonas
    Set objPersona = New COMDPersona.DCOMPersonas
    
    Dim oMov As COMDMov.DCOMMov
    Set oMov = New COMDMov.DCOMMov
    
    nMontoComis = val(lblMontoCom.Caption)
    lsClaveSeg = Trim(Me.txtClaveNew.Text)
    
    If lsClaveSeg = "" Then
        MsgBox "El Remitente no Ingreso Clave de Seguridad "
        Exit Sub
    End If
    
    If GClaveGiro <> Me.txtClaveAnt.Text Then
       MsgBox "Clave de seguridad no valida. Ingrese la clave correcta", vbCritical, "Aviso"
       Exit Sub
    End If
    
    'If txtNombre.Text = "" Or txtDNI.Text = "" Or txtReferencia = "" Or txtClaveNew.Text = "" Then
    If TxtNombre.Text = "" Or txtDNI.Text = "" Or txtReferencia = "" Or txtClaveNew.Text = "" Or txtBuscaPersona.Text = "" Then 'RECO20140722
        MsgBox "Los datos no pueden ser vacios.", vbCritical, "Aviso"
        Exit Sub
    End If
    
    If Me.TxtNombre.Text <> "" Then
        'If (Right(cboAgencia.Text, 2) = Right(Me.lblAgenciaDes.Caption, 2)) Then
            'MsgBox "La nueva agencia destino no puede ser la agencia destino anterior", vbCritical, "Aviso"
            'Exit Sub
        'End If
    End If
    
    
    
    If MsgBox("¿Desea Grabar la Operación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrGraba
    Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim sMovNro As String, sCuentaGiro As String, sPersLavDinero As String
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    
    Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion
        
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    
    Dim psAgeDestino  As String
    sAgenciaDest = Right(cboagencia.Text, 2)
    nTpoClienteAct = IIf(Me.chkClienteCMAC.value = 1, 1, 2)
    
    Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
    psAgeDestino = Left(cboagencia.Text, 30)
    clsServ.ServGiroCambioDestino Me.txtCuenta.NroCuenta, sMovNro, val(Me.lblMonto.Caption), nMontoComis, nTpoClienteAnt, nTpoClienteAct, sAgenciaDest, Right(Me.lblAgenciaDes.Caption, 2), Me.txtBuscaPersona.Text, txtDNI.Text, sPersCodAnt, sNumDOIAnt, Me.TxtNombre.Text, txtReferencia.Text, Me.txtClaveNew.Text, lsBoleta, sRemitenteNomb, Me.cboagencia.Text
    
    Set clsServ = Nothing
        
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
Exit Sub
ErrGraba:
    MsgBox err.Description, vbExclamation, "Error"
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CargaAgencias
End Sub

Private Sub txtBuscaPersona_EmiteDatos()
    Dim fnPersPersoneriaGen As Integer
    Dim fcPersCod As String
    If Trim(txtBuscaPersona.sPersNroDoc) = "" Then
        MsgBox "Esta persona no tiene un documento de identidad ingresado." & vbCrLf & " Por favor actualice su información.", vbOKOnly + vbInformation, "Atención"
        txtBuscaPersona.Text = ""
        Exit Sub
    End If
    If txtBuscaPersona.Text <> "" And txtBuscaPersona.sPersNroDoc <> "" Then
        If sRemitente = txtBuscaPersona.Text Then
            MsgBox "El destinatario no puede ser igual al remitente", vbCritical, "Aviso"
            Exit Sub
        End If
        'txtReferencia.Text = txtBuscaPersona.sPersDireccion
        TxtNombre.Text = txtBuscaPersona.psDescripcion
        sCodRemitente = txtBuscaPersona.Text
        SNroIdentidad = txtBuscaPersona.sPersNroDoc
        txtDNI.Text = txtBuscaPersona.sPersNroDoc
        Dim dlsMant As COMDCaptaGenerales.DCOMCaptaGenerales
        Set dlsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
        fnPersPersoneriaGen = IIf(txtBuscaPersona.PersPersoneria > 1, 2, 1)
        fcPersCod = Trim(txtBuscaPersona.psCodigoPersona)
        If dlsMant.GetNroOPeradoras(gsCodAge) > 1 Then
            If txtBuscaPersona.psCodigoPersona = gsCodPersUser Then
                MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                Unload Me
                Exit Sub
            End If
        End If
        Set dlsMant = Nothing
    End If
End Sub
Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCuenta As String, sMoneda As String
        sCuenta = txtCuenta.NroCuenta
        CargaDatosGiro sCuenta
        'If fcAgeCodDest <> gsCodAge Then
        '    MsgBox "No se puede realizar la operación, Giro pertenece a otra agencia", vbCritical, "Aviso"
        '    Exit Sub
        'End If
        Me.frmDest.Enabled = True
        Me.frmClave.Enabled = True
        Me.frmAgencia.Enabled = True
        Me.cmdAceptar.Enabled = True
    End If
End Sub
Private Sub CargaAgencias()
    Dim clsAge As COMDConstSistema.DCOMGeneral
    Dim rs As ADODB.Recordset
    Dim i As Long
    Set rs = New ADODB.Recordset
    Set clsAge = New COMDConstSistema.DCOMGeneral
    Set rs = clsAge.GetNombreAgencias(gsCodAge)
    Do While Not rs.EOF
        cboagencia.AddItem rs("cAgeDescripcion") & Space(50) & rs("cAgeCod")
        rs.MoveNext
    Loop
    cboagencia.ListIndex = 0
    rs.Close
    Set clsAge = Nothing
    Set rs = Nothing
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
        lblAgenciaDes = Trim(rsGiro("cAgencia")) & Space(200) & rsGiro("cAgenciaDest")
        lblMonto = Format$(rsGiro("nSaldo"), "#,##0.00")
        lblTpoGiro = Trim(rsGiro("cTipo"))
        lblFecAper = Format$(rsGiro("dPrdEstado"), "dd mmm yyyy")
        sRemitente = Trim(rsGiro("cPersCod"))
        nTpoClienteAnt = IIf(rsGiro("cNumDocDesti") = "", 1, 2)
        sPersCodAnt = rsGiro("cCodDest")
        sNumDOIAnt = rsGiro("cNumDocDesti")
        sRemitenteNomb = rsGiro("cRemitente")
        fcAgeCodDest = Trim(rsGiro("cAgenciaDest"))
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
                lblMoneda.Caption = "US$"
            Else
                'lblMonto.BackColor = &HFFFFFF
                lblMoneda.Caption = "S/."
            End If
            'SendKeys "{Enter}"
        End If
        If nTpoClienteAnt = 1 Then
            Call HabilitarControles(False)
        Else
            Call HabilitarControles(True)
        End If
    Else
        MsgBox "Número de Giro no encontrado o Cancelado.", vbInformation, "SICMACM - Aviso"
        txtCuenta.Age = ""
        txtCuenta.Cuenta = ""
        txtCuenta.SetFocusAge
        sRemitente = ""
    End If
    
    Set rsGiroCom = clsGiro.RecuperaValorComisionTarGiro(1)
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
Public Sub HabilitarControles(ByVal bHabilitar As Boolean)
    TxtNombre.Enabled = bHabilitar
    txtDNI.Enabled = bHabilitar
    'txtReferencia.Enabled = bHabilitar
    Me.txtBuscaPersona.Enabled = Not bHabilitar
    Me.txtBuscaPersona.Visible = Not bHabilitar
End Sub
Public Sub LimpiarFormulario()
    txtCuenta.Age = ""
    txtCuenta.Cuenta = ""
    lblFecAper.Caption = ""
    lblTpoGiro.Caption = ""
    lblAgenciaDes.Caption = ""
    lblMonto.Caption = ""
    txtBuscaPersona.Text = ""
    TxtNombre.Text = ""
    txtDNI.Text = ""
    txtReferencia.Text = ""
    txtClaveAnt.Text = ""
    txtClaveNew.Text = ""
    cboagencia.ListIndex = 0
    lblMoneda.Caption = ""
    lblMontoCom.Caption = ""
End Sub
