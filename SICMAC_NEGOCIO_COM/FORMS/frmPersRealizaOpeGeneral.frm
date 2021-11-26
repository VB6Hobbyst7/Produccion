VERSION 5.00
Begin VB.Form frmPersRealizaOpeGeneral 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Usuarios Relacionados a la Transacción"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10665
   Icon            =   "frmPersRealizaOpeGeneral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOrigen 
      Height          =   330
      Left            =   1680
      MaxLength       =   300
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   2640
      Width           =   6495
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame fraDatosOpe 
      Caption         =   "Intervinientes"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   10455
      Begin VB.ComboBox cmdOpeEfectivo 
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1920
         Width           =   1095
      End
      Begin VB.ComboBox cmdPresencia 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtNombreBeneficia 
         Height          =   330
         Left            =   4440
         MaxLength       =   200
         TabIndex        =   10
         Top             =   1440
         Width           =   5535
      End
      Begin VB.ComboBox cboTipoDOIBeneficia 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtNDOIBeneficia 
         Height          =   330
         Left            =   2760
         TabIndex        =   9
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdBuscarPersBeneficia 
         Caption         =   "..."
         Height          =   330
         Left            =   9960
         TabIndex        =   11
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtNombreRealiza 
         Height          =   330
         Left            =   4440
         MaxLength       =   200
         TabIndex        =   6
         Top             =   960
         Width           =   5535
      End
      Begin VB.ComboBox cboTipoDOIRealiza 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtNDOIRealiza 
         Height          =   330
         Left            =   2760
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdBuscarPersRealiza 
         Caption         =   "..."
         Height          =   330
         Left            =   9960
         TabIndex        =   7
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdBuscarPersOrdena 
         Caption         =   "..."
         Height          =   330
         Left            =   9960
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtNDOIOrdena 
         Height          =   330
         Left            =   2760
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cboTipoDOIOrdena 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtNombreOrdena 
         Height          =   330
         Left            =   4440
         MaxLength       =   200
         TabIndex        =   2
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "presente?"
         Height          =   195
         Left            =   720
         TabIndex        =   27
         Top             =   2040
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "La Operacion es en Efectivo:"
         Height          =   195
         Left            =   2760
         TabIndex        =   26
         Top             =   1920
         Width           =   2070
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Usuario está"
         Height          =   195
         Left            =   600
         TabIndex        =   25
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label lblBeneficia 
         AutoSize        =   -1  'True
         Caption         =   "El que se beneficia:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblRealiza 
         AutoSize        =   -1  'True
         Caption         =   "El que realiza:"
         Height          =   195
         Left            =   480
         TabIndex        =   23
         Top             =   960
         Width           =   990
      End
      Begin VB.Label lblOrdena 
         AutoSize        =   -1  'True
         Caption         =   "El que Ordena:"
         Height          =   195
         Left            =   360
         TabIndex        =   21
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº DOI:"
         Height          =   195
         Left            =   3240
         TabIndex        =   20
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo DOI:"
         Height          =   195
         Left            =   1680
         TabIndex        =   19
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   4680
         TabIndex        =   18
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Label lblOrigen 
      AutoSize        =   -1  'True
      Caption         =   "Origen de los Fondos:"
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   1545
   End
End
Attribute VB_Name = "frmPersRealizaOpeGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre :frmPersRealizaOpeGeneral
'** Descripción : Nos permite identificar los intervinientes de una operacion dentro de los parametros internos
'**               creado segun TI-ERS0005-2013
'** Creación : WIOR 20130301
'**********************************************************************************************

Option Explicit
'ORDENA
Private lsPersCodOrd As String
Private lsPersNombreOrd As String
Private lsPersDOIOrd As String
Private lsPersTipoDOIOrd As String
Private lsTipoClienteOrd As String
'REALIZA
Private lsPersCodRea As String
Private lsPersNombreRea As String
Private lsPersDOIRea As String
Private lsPersTipoDOIRea As String
Private lsTipoClienteRea As String
'BENEFICIA
Private lsPersCodBen As String
Private lsPersNombreBen As String
Private lsPersDOIBen As String
Private lsPersTipoDOIBen As String
Private lsTipoClienteBen As String
'OPERACION
Private lsPersRegistrar As Boolean
Private lsTipoOperacion As String
Private lsOrigenFondos As String
Private lnUsuPresente As Integer
Private lnOpeEfectivo As Integer
Dim i As Integer
'PERSONA ORDENA
Property Let PersCodOrd(pPersCodOrd As String)
   lsPersCodOrd = pPersCodOrd
End Property
Property Get PersCodOrd() As String
    PersCodOrd = lsPersCodOrd
End Property
Property Let PersNombreOrd(pPersNombreOrd As String)
   lsPersNombreOrd = pPersNombreOrd
End Property
Property Get PersNombreOrd() As String
    PersNombreOrd = lsPersNombreOrd
End Property
Property Let PersDOIOrd(pPersDOIOrd As String)
   lsPersDOIOrd = pPersDOIOrd
End Property
Property Get PersDOIOrd() As String
    PersDOIOrd = lsPersDOIOrd
End Property
Property Let PersTipoDOIOrd(pPersTipoDOIOrd As String)
   lsPersTipoDOIOrd = pPersTipoDOIOrd
End Property
Property Get PersTipoDOIOrd() As String
    PersTipoDOIOrd = lsPersTipoDOIOrd
End Property
Property Let PersTipoClienteOrd(pPersTipoClienteOrd As String)
   lsTipoClienteOrd = pPersTipoClienteOrd
End Property
Property Get PersTipoClienteOrd() As String
    PersTipoClienteOrd = lsTipoClienteOrd
End Property
'PERSONA REALIZA
Property Let PersCodRea(pPersCodRea As String)
   lsPersCodRea = pPersCodRea
End Property
Property Get PersCodRea() As String
    PersCodRea = lsPersCodRea
End Property
Property Let PersNombreRea(pPersNombreRea As String)
   lsPersNombreRea = pPersNombreRea
End Property
Property Get PersNombreRea() As String
    PersNombreRea = lsPersNombreRea
End Property
Property Let PersDOIRea(pPersDOIRea As String)
   lsPersDOIRea = pPersDOIRea
End Property
Property Get PersDOIRea() As String
    PersDOIRea = lsPersDOIRea
End Property
Property Let PersTipoDOIRea(pPersTipoDOIRea As String)
   lsPersTipoDOIRea = pPersTipoDOIRea
End Property
Property Get PersTipoDOIRea() As String
    PersTipoDOIRea = lsPersTipoDOIRea
End Property
Property Let PersTipoClienteRea(pPersTipoClienteRea As String)
   lsTipoClienteRea = pPersTipoClienteRea
End Property
Property Get PersTipoClienteRea() As String
    PersTipoClienteRea = lsTipoClienteRea
End Property
'PERSONA BENEFICIA
Property Let PersCodBen(pPersCodBen As String)
   lsPersCodBen = pPersCodBen
End Property
Property Get PersCodBen() As String
    PersCodBen = lsPersCodBen
End Property
Property Let PersNombreBen(pPersNombreBen As String)
   lsPersNombreBen = pPersNombreBen
End Property
Property Get PersNombreBen() As String
    PersNombreBen = lsPersNombreBen
End Property
Property Let PersDOIBen(pPersDOIBen As String)
   lsPersDOIBen = pPersDOIBen
End Property
Property Get PersDOIBen() As String
    PersDOIBen = lsPersDOIBen
End Property
Property Let PersTipoDOIBen(pPersTipoDOIBen As String)
   lsPersTipoDOIBen = pPersTipoDOIBen
End Property
Property Get PersTipoDOIBen() As String
    PersTipoDOIBen = lsPersTipoDOIBen
End Property
Property Let PersTipoClienteBen(pPersTipoClienteBen As String)
   lsTipoClienteBen = pPersTipoClienteBen
End Property
Property Get PersTipoClienteBen() As String
    PersTipoClienteBen = lsTipoClienteBen
End Property
'OPERACION
Property Let PersRegistrar(pPersReg As Boolean)
   lsPersRegistrar = pPersReg
End Property
Property Get PersRegistrar() As Boolean
    PersRegistrar = lsPersRegistrar
End Property

Property Let TipoOperacion(pTpoOpe As String)
   lsTipoOperacion = pTpoOpe
End Property
Property Get TipoOperacion() As String
    TipoOperacion = lsTipoOperacion
End Property
Property Let Origen(psOrigen As String)
   lsOrigenFondos = psOrigen
End Property
Property Get Origen() As String
    Origen = lsOrigenFondos
End Property
Property Let UsuPresente(pnUsuPresente As Integer)
   lnUsuPresente = pnUsuPresente
End Property
Property Get UsuPresente() As Integer
    UsuPresente = lnUsuPresente
End Property
Property Let OpeEfectivo(pnOpeEfectivo As Integer)
   lnOpeEfectivo = pnOpeEfectivo
End Property
Property Get OpeEfectivo() As Integer
    OpeEfectivo = lnOpeEfectivo
End Property

Private Sub cboTipoDOIBeneficia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNDOIBeneficia.SetFocus
End If
End Sub

Private Sub cboTipoDOIOrdena_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNDOIOrdena.SetFocus
End If
End Sub

Private Sub cboTipoDOIRealiza_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNDOIRealiza.SetFocus
End If
End Sub

Private Sub cmdAceptar_Click()
If ValidaDatos Then
    If MsgBox("Esta seguro de grabar los datos ", vbInformation + vbYesNo, "Aviso") = vbYes Then
    'PERSONA ORDENA
    If lsTipoClienteOrd = 0 Then
        lsPersCodOrd = ""
        lsPersNombreOrd = Trim(txtNombreOrdena.Text)
        lsPersDOIOrd = Trim(txtNDOIOrdena.Text)
        lsPersTipoDOIOrd = Trim(Right(cboTipoDOIOrdena.Text, 5))
    End If
    'PERSONA REALIZA
    If lsTipoClienteRea = 0 Then
        lsPersCodRea = ""
        lsPersNombreRea = Trim(txtNombreRealiza.Text)
        lsPersDOIRea = Trim(txtNDOIRealiza.Text)
        lsPersTipoDOIRea = Trim(Right(cboTipoDOIRealiza.Text, 5))
    End If
    'PERSONA BENEFICIA
    If lsTipoClienteBen = 0 Then
        lsPersCodBen = ""
        lsPersNombreBen = Trim(txtNombreBeneficia.Text)
        lsPersDOIBen = Trim(txtNDOIBeneficia.Text)
        lsPersTipoDOIBen = Trim(Right(cboTipoDOIBeneficia.Text, 5))
    End If
    
    If Me.txtOrigen.Visible Then
        lsOrigenFondos = Trim(txtOrigen.Text)
    Else
        lsOrigenFondos = ""
    End If
    
    lnUsuPresente = Trim(Right(cmdPresencia.Text, 5))
    lnOpeEfectivo = Trim(Right(cmdOpeEfectivo.Text, 5))

    lsPersRegistrar = True
    Unload Me
    End If
End If
End Sub

Private Sub cmdBuscarPersBeneficia_Click()
Dim oPers As COMDPersona.UCOMPersona
Dim oPersona As UPersona_Cli
Dim nTipoDOI As Integer
Dim sTipoDOI As String
Dim sNumeroDOI As String

Call HabilitaControles(True, 1, 3)

    
    Set oPers = frmBuscaPersona.Inicio
    Set oPersona = New UPersona_Cli
    If Not oPers Is Nothing Then
        oPersona.RecuperaPersona (oPers.sPersCod)
    Else
        Call HabilitaControles(True, 1, 3)
        Exit Sub
    End If
    
    If Not oPersona Is Nothing Then
        Call oPersona.ObtenerDatosDocumentoxPos(0, nTipoDOI, sTipoDOI, sNumeroDOI)
        txtNombreBeneficia.Text = PstaNombre(oPersona.NombreCompleto, True)
        cboTipoDOIBeneficia.ListIndex = IndiceListaCombo(cboTipoDOIBeneficia, nTipoDOI)
        txtNDOIBeneficia.Text = sNumeroDOI
        Call HabilitaControles(False, , 3)
        lsPersCodBen = oPers.sPersCod
        lsPersNombreBen = PstaNombre(oPersona.NombreCompleto, True)
        lsPersDOIBen = sNumeroDOI
        lsPersTipoDOIBen = nTipoDOI
        lsTipoClienteBen = 1
        cmdAceptar.SetFocus
    Else
        Call HabilitaControles(True, 1, 3)
        Exit Sub
    End If
    Set oPersona = Nothing
End Sub

Private Sub cmdBuscarPersOrdena_Click()
Dim oPers As COMDPersona.UCOMPersona
Dim oPersona As UPersona_Cli
Dim nTipoDOI As Integer
Dim sTipoDOI As String
Dim sNumeroDOI As String

Call HabilitaControles(True, 1, 1)

    
    Set oPers = frmBuscaPersona.Inicio
    Set oPersona = New UPersona_Cli
    If Not oPers Is Nothing Then
        oPersona.RecuperaPersona (oPers.sPersCod)
    Else
        Call HabilitaControles(True, 1, 1)
        Exit Sub
    End If
    
    If Not oPersona Is Nothing Then
        Call oPersona.ObtenerDatosDocumentoxPos(0, nTipoDOI, sTipoDOI, sNumeroDOI)
        txtNombreOrdena.Text = PstaNombre(oPersona.NombreCompleto, True)
        cboTipoDOIOrdena.ListIndex = IndiceListaCombo(cboTipoDOIOrdena, nTipoDOI)
        txtNDOIOrdena.Text = sNumeroDOI
        Call HabilitaControles(False, , 1)
        lsPersCodOrd = oPers.sPersCod
        lsPersNombreOrd = PstaNombre(oPersona.NombreCompleto, True)
        lsPersDOIOrd = sNumeroDOI
        lsPersTipoDOIOrd = nTipoDOI
        lsTipoClienteOrd = 1
        
        If Trim(Right(cmdPresencia.Text, 5)) = "2" Then
            cmdBuscarPersBeneficia.SetFocus
        Else
            cmdBuscarPersRealiza.SetFocus
        End If
    Else
        Call HabilitaControles(True, 1, 1)
        Exit Sub
    End If
    Set oPersona = Nothing
End Sub

Private Sub cmdBuscarPersRealiza_Click()
Dim oPers As COMDPersona.UCOMPersona
Dim oPersona As UPersona_Cli
Dim nTipoDOI As Integer
Dim sTipoDOI As String
Dim sNumeroDOI As String

Call HabilitaControles(True, 1, 2)

    
    Set oPers = frmBuscaPersona.Inicio
    Set oPersona = New UPersona_Cli
    If Not oPers Is Nothing Then
        oPersona.RecuperaPersona (oPers.sPersCod)
    Else
        Call HabilitaControles(True, 1, 2)
        Exit Sub
    End If
    
    If Not oPersona Is Nothing Then
    
        If oPersona.Personeria = "1" Then
            Call oPersona.ObtenerDatosDocumentoxPos(0, nTipoDOI, sTipoDOI, sNumeroDOI)
            txtNombreRealiza.Text = PstaNombre(oPersona.NombreCompleto, True)
            cboTipoDOIRealiza.ListIndex = IndiceListaCombo(cboTipoDOIRealiza, nTipoDOI)
            txtNDOIRealiza.Text = sNumeroDOI
            Call HabilitaControles(False, , 2)
            lsPersCodRea = oPers.sPersCod
            lsPersNombreRea = PstaNombre(oPersona.NombreCompleto, True)
            lsPersDOIRea = sNumeroDOI
            lsPersTipoDOIRea = nTipoDOI
            lsTipoClienteRea = 1
            cmdBuscarPersBeneficia.SetFocus
        Else
            MsgBox "No es una Persona Natural", vbInformation, "Aviso"
            Call HabilitaControles(True, 1, 2)
            lsTipoClienteRea = 0
            cmdBuscarPersRealiza.SetFocus
            Exit Sub
        End If

    Else
        Call HabilitaControles(True, 1, 2)
        Exit Sub
    End If
    Set oPersona = Nothing
End Sub

Private Sub cmdCancelar_Click()
    lsPersRegistrar = False
    Unload Me
End Sub

Private Sub CargaControles()
    lsPersRegistrar = False
    lsTipoClienteOrd = 0
    lsTipoClienteBen = 0
    lsTipoClienteRea = 0
    'TIPO DOI
    'ORDENA
    Call CargaComboConstante(gPersIdTipo, cboTipoDOIOrdena)
    Call CambiaTamañoCombo(cboTipoDOIOrdena, 150)
    'REALIZA
    Call CargaComboConstante(gPersIdTipo, cboTipoDOIRealiza)
    Call CambiaTamañoCombo(cboTipoDOIRealiza, 150)
    'BENEFICIA
    Call CargaComboConstante(gPersIdTipo, cboTipoDOIBeneficia)
    Call CambiaTamañoCombo(cboTipoDOIBeneficia, 150)
    
    'USUARIO PRESENTE
    Call CargaComboConstante(4046, cmdPresencia)
    'OPERACION EN EFECTIVO
    Call CargaComboConstante(4046, cmdOpeEfectivo)
    
    Call HabilitaControles(True, 1, 4)
End Sub

Private Sub CmdLimpiar_Click()
Call HabilitaControles(True, 1)
End Sub
Public Sub Inicia(ByVal psDesc As String, ByVal pnOpeCod As String, Optional ByVal pnCliReforzado As Integer = 0)
Me.Caption = Me.Caption & " - " & psDesc
lsTipoOperacion = pnOpeCod
Call CargaControles

If pnCliReforzado = 0 Then
    lblOrigen.Visible = False
    txtOrigen.Visible = False
ElseIf pnCliReforzado = 1 Then
    lblOrigen.Visible = True
    txtOrigen.Visible = True
End If

Me.Show 1
End Sub
Private Sub HabilitaControles(ByVal pbHabilita As Boolean, Optional ByVal pnLimpia As Integer = 0, Optional ByVal pnTipoInter As Integer)
If pnTipoInter = 1 Then
    txtNombreOrdena.Enabled = pbHabilita
    txtNDOIOrdena.Enabled = pbHabilita
    cboTipoDOIOrdena.Enabled = pbHabilita
    If pnLimpia <> 0 Then
        txtNombreOrdena.Text = ""
        txtNDOIOrdena.Text = ""
        cboTipoDOIOrdena.ListIndex = IndiceListaCombo(cboTipoDOIOrdena, 0)
        lsTipoClienteOrd = 0
        txtNombreOrdena.SetFocus
    End If
ElseIf pnTipoInter = 2 Then
    txtNombreRealiza.Enabled = pbHabilita
    txtNDOIRealiza.Enabled = pbHabilita
    cboTipoDOIRealiza.Enabled = pbHabilita
    If pnLimpia <> 0 Then
        txtNombreRealiza.Text = ""
        txtNDOIRealiza.Text = ""
        cboTipoDOIRealiza.ListIndex = IndiceListaCombo(cboTipoDOIRealiza, 0)
        lsTipoClienteRea = 0
        txtNombreRealiza.SetFocus
    End If
ElseIf pnTipoInter = 3 Then
    txtNombreBeneficia.Enabled = pbHabilita
    txtNDOIBeneficia.Enabled = pbHabilita
    cboTipoDOIBeneficia.Enabled = pbHabilita
    If pnLimpia <> 0 Then
        txtNombreBeneficia.Text = ""
        txtNDOIBeneficia.Text = ""
        cboTipoDOIBeneficia.ListIndex = IndiceListaCombo(cboTipoDOIBeneficia, 0)
        lsTipoClienteBen = 0
        txtNombreBeneficia.SetFocus
    End If
Else
    txtNombreOrdena.Enabled = pbHabilita
    txtNDOIOrdena.Enabled = pbHabilita
    cboTipoDOIOrdena.Enabled = pbHabilita
    
    txtNombreRealiza.Enabled = pbHabilita
    txtNDOIRealiza.Enabled = pbHabilita
    cboTipoDOIRealiza.Enabled = pbHabilita
    
    txtNombreBeneficia.Enabled = pbHabilita
    txtNDOIBeneficia.Enabled = pbHabilita
    cboTipoDOIBeneficia.Enabled = pbHabilita
    
    
    If pnLimpia <> 0 Then
        txtNombreOrdena.Text = ""
        txtNDOIOrdena.Text = ""
        cboTipoDOIOrdena.ListIndex = IndiceListaCombo(cboTipoDOIOrdena, 0)
        lsTipoClienteOrd = 0
    
        txtNombreRealiza.Text = ""
        txtNDOIRealiza.Text = ""
        cboTipoDOIRealiza.ListIndex = IndiceListaCombo(cboTipoDOIRealiza, 0)
        lsTipoClienteRea = 0
        
        txtNombreBeneficia.Text = ""
        txtNDOIBeneficia.Text = ""
        cboTipoDOIBeneficia.ListIndex = IndiceListaCombo(cboTipoDOIBeneficia, 0)
        lsTipoClienteBen = 0
        
        lsPersCodOrd = ""
        lsPersNombreOrd = ""
        lsPersDOIOrd = ""
        lsPersTipoDOIOrd = 0
        
        lsPersCodRea = ""
        lsPersNombreRea = ""
        lsPersDOIRea = ""
        lsPersTipoDOIRea = 0

        lsPersCodBen = ""
        lsPersNombreBen = ""
        lsPersDOIBen = ""
        lsPersTipoDOIBen = 0
    
        lsOrigenFondos = ""
        lnUsuPresente = 0
        lnOpeEfectivo = 0
    End If
End If

End Sub
Private Function ValidaDatos() As Boolean
Dim ClsPersona As COMDPersona.DCOMPersonas
Set ClsPersona = New COMDPersona.DCOMPersonas
Dim R As ADODB.Recordset

If Trim(cmdPresencia.Text) = "" Then
    MsgBox "Favor de seleccionar si el Usuario Esta Presente o No.", vbInformation, "Aviso"
    cmdPresencia.SetFocus
    ValidaDatos = False
    Exit Function
End If

If Trim(cmdOpeEfectivo.Text) = "" Then
    MsgBox "Favor de seleccionar si la operacion es en efectivo o No.", vbInformation, "Aviso"
    cmdOpeEfectivo.SetFocus
    ValidaDatos = False
    Exit Function
End If

If lsTipoClienteOrd = 0 Then
    'PERSONA ORDENA
        If Trim(txtNombreOrdena.Text) = "" Then
            MsgBox "Ingrese el Nombre de la Persona que Ordena la Operación.", vbInformation, "Aviso"
            txtNombreOrdena.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If Trim(Me.cboTipoDOIOrdena.Text) = "" Then
            MsgBox "Seleccione el Tipo de DOI de la Persona que Ordena la Operación.", vbInformation, "Aviso"
            cboTipoDOIOrdena.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If Trim(txtNDOIOrdena.Text) = "" Then
            MsgBox "Ingrese el Nro. de DOI de la Persona que Ordena la Operación.", vbInformation, "Aviso"
            cboTipoDOIOrdena.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        
        'Verfica Longitud de DOI
        If Trim(cboTipoDOIOrdena.Text) <> "" Then
            If CInt(Trim(Right(cboTipoDOIOrdena.Text, 5))) = gPersIdDNI Then
                If Len(Trim(txtNDOIOrdena.Text)) <> gnNroDigitosDNI Then
                    MsgBox "DNI de la Persona que Ordena No es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
                    txtNDOIOrdena.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
            End If
            If CInt(Trim(Right(cboTipoDOIOrdena.Text, 5))) = gPersIdRUC Then
                If Len(Trim(txtNDOIOrdena.Text)) <> gnNroDigitosRUC Then
                    MsgBox "RUC de la Persona que Ordena No es de " & gnNroDigitosRUC & " digitos", vbInformation, "Aviso"
                    txtNDOIOrdena.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
            End If
        End If
       
        Set R = ClsPersona.BuscaCliente(txtNDOIOrdena.Text, 3)
        If R.RecordCount > 0 Then
            MsgBox "Nro. de DOI " & txtNDOIOrdena.Text & " de la persona que Ordena ya existe el la base de datos, Favor de Buscar a la persona como cliente CMACMaynas.", vbInformation, "Aviso"
            txtNDOIOrdena.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        Set R = Nothing
End If

If lsTipoClienteRea = 0 Then
    'PERSONA REALIZA
    If Trim(Right(cmdPresencia.Text, 5)) <> "2" Then
        If Trim(txtNombreRealiza.Text) = "" Then
            MsgBox "Ingrese el Nombre de la Persona que Realiza la Operación.", vbInformation, "Aviso"
            txtNombreRealiza.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If Trim(Me.cboTipoDOIRealiza.Text) = "" Then
            MsgBox "Seleccione el Tipo de DOI de la Persona que Realiza la Operación.", vbInformation, "Aviso"
            cboTipoDOIRealiza.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If Trim(txtNDOIRealiza.Text) = "" Then
            MsgBox "Ingrese el Nro. de DOI de la Persona que Realiza la Operación.", vbInformation, "Aviso"
            cboTipoDOIRealiza.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        
        'Verfica Longitud de DOI
        If Trim(cboTipoDOIRealiza.Text) <> "" Then
            If CInt(Trim(Right(cboTipoDOIRealiza.Text, 5))) = gPersIdDNI Then
                If Len(Trim(txtNDOIRealiza.Text)) <> gnNroDigitosDNI Then
                    MsgBox "DNI de la Persona que Realiza No es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
                    txtNDOIRealiza.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
            End If
            If CInt(Trim(Right(cboTipoDOIRealiza.Text, 5))) = gPersIdRUC Then
                If Len(Trim(txtNDOIRealiza.Text)) <> gnNroDigitosRUC Then
                    MsgBox "RUC de la Persona que Realiza No es de " & gnNroDigitosRUC & " digitos", vbInformation, "Aviso"
                    txtNDOIRealiza.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
            End If
        End If
       
        Set R = ClsPersona.BuscaCliente(txtNDOIRealiza.Text, 3)
        If R.RecordCount > 0 Then
            MsgBox "Nro. de DOI " & txtNDOIRealiza.Text & " de la Persona que Realiza ya existe el la base de datos, Favor de Buscar a la persona como cliente CMACMaynas.", vbInformation, "Aviso"
            txtNDOIRealiza.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        Set R = Nothing
    End If
    
End If
    

If lsTipoClienteBen = 0 Then
    'PERSONA BENEFICIA
        If Trim(txtNombreBeneficia.Text) = "" Then
            MsgBox "Ingrese el Nombre de la Persona que se Beneficia con la Operación.", vbInformation, "Aviso"
            txtNombreBeneficia.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If Trim(Me.cboTipoDOIBeneficia.Text) = "" Then
            MsgBox "Seleccione el Tipo de DOI de la Persona que se Beneficia con la Operación.", vbInformation, "Aviso"
            cboTipoDOIBeneficia.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If Trim(txtNDOIBeneficia.Text) = "" Then
            MsgBox "Ingrese el Nro. de DOI de la Persona que se Beneficia con la Operación.", vbInformation, "Aviso"
            cboTipoDOIBeneficia.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        
        'Verfica Longitud de DOI
        If Trim(cboTipoDOIBeneficia.Text) <> "" Then
            If CInt(Trim(Right(cboTipoDOIBeneficia.Text, 5))) = gPersIdDNI Then
                If Len(Trim(txtNDOIBeneficia.Text)) <> gnNroDigitosDNI Then
                    MsgBox "DNI de la Persona que se Beneficia No es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
                    txtNDOIBeneficia.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
            End If
            If CInt(Trim(Right(cboTipoDOIBeneficia.Text, 5))) = gPersIdRUC Then
                If Len(Trim(txtNDOIBeneficia.Text)) <> gnNroDigitosRUC Then
                    MsgBox "RUC de la Persona que se Beneficia No es de " & gnNroDigitosRUC & " digitos", vbInformation, "Aviso"
                    txtNDOIBeneficia.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
            End If
        End If
       
        Set R = ClsPersona.BuscaCliente(txtNDOIBeneficia.Text, 3)
        If R.RecordCount > 0 Then
            MsgBox "Nro. de DOI " & txtNDOIBeneficia.Text & " de la Persona que se Beneficia ya existe el la base de datos, Favor de Buscar a la persona como cliente CMACMaynas.", vbInformation, "Aviso"
            txtNDOIBeneficia.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        Set R = Nothing
End If

    If txtOrigen.Visible Then
        If Trim(txtOrigen.Text) = "" Then
            MsgBox "Ingrese el Origen de los Fondos", vbInformation, "Aviso"
            txtOrigen.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        
        If Len(txtOrigen.Text) >= 300 Then
            MsgBox "El texto de Origen no debe superar 300 caracteres, Favor Resumir.", vbInformation, "Aviso"
            txtOrigen.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If


ValidaDatos = True
End Function

Public Sub InsertaPersonasOperacion(ByVal pnNroMov As Long, ByVal psctacod As String, Optional ByVal pnCondicion As Integer = 0)
    Dim ListaIntervinientes() As PersOpeGeneral
    Dim oNPersona As COMNPersona.NCOMPersona
    
    If Trim(lsPersDOIOrd) <> "" Then
        ReDim Preserve ListaIntervinientes(0)
        ListaIntervinientes(0).PersCod = lsPersCodOrd
        ListaIntervinientes(0).TpoCliente = lsTipoClienteOrd
        ListaIntervinientes(0).TpoDOI = lsPersTipoDOIOrd
        ListaIntervinientes(0).NroDOI = lsPersDOIOrd
        ListaIntervinientes(0).Nombre = lsPersNombreOrd
        ListaIntervinientes(0).TipoRelacion = OrdenaTransaccion
    End If
    
    If Trim(lsPersDOIBen) <> "" Then
        ReDim Preserve ListaIntervinientes(1)
        ListaIntervinientes(1).PersCod = lsPersCodBen
        ListaIntervinientes(1).TpoCliente = lsTipoClienteBen
        ListaIntervinientes(1).TpoDOI = lsPersTipoDOIBen
        ListaIntervinientes(1).NroDOI = lsPersDOIBen
        ListaIntervinientes(1).Nombre = lsPersNombreBen
        ListaIntervinientes(1).TipoRelacion = BeneficiaTransaccion
    End If
    
    If lnUsuPresente = 1 Then
        If Trim(lsPersDOIRea) <> "" Then
            ReDim Preserve ListaIntervinientes(2)
            ListaIntervinientes(2).PersCod = lsPersCodRea
            ListaIntervinientes(2).TpoCliente = lsTipoClienteRea
            ListaIntervinientes(2).TpoDOI = lsPersTipoDOIRea
            ListaIntervinientes(2).NroDOI = lsPersDOIRea
            ListaIntervinientes(2).Nombre = lsPersNombreRea
            ListaIntervinientes(2).TipoRelacion = RealizaTransaccion
        End If
    End If
    
    Set oNPersona = New COMNPersona.NCOMPersona
    Call oNPersona.InsertaPersonasOperacionGen(pnNroMov, psctacod, lnUsuPresente, lnOpeEfectivo, lsOrigenFondos, lsTipoOperacion, pnCondicion, ListaIntervinientes)
    Set oNPersona = Nothing
    
    Call HabilitaControles(True, 1, 4)
End Sub
Private Sub cmdPresencia_Click()
If Trim(Right(cmdPresencia.Text, 5)) = "2" Then
    Call HabilitaControles(True, 1, 2)
    cboTipoDOIRealiza.ListIndex = -1
    cboTipoDOIRealiza.Enabled = False
    txtNDOIRealiza.Enabled = False
    txtNombreRealiza.Enabled = False
    cmdBuscarPersRealiza.Enabled = False
Else
    If lsTipoClienteRea = 0 Then
    cboTipoDOIRealiza.Enabled = True
    txtNDOIRealiza.Enabled = True
    txtNombreRealiza.Enabled = True
    cmdBuscarPersRealiza.Enabled = True
    End If
End If
End Sub

Private Sub txtNDOIBeneficia_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtNDOIBeneficia, KeyAscii, 30)
    If KeyAscii = 13 Then
        txtNombreBeneficia.SetFocus
    End If
End Sub

Private Sub txtNDOIOrdena_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtNDOIOrdena, KeyAscii, 30)
    If KeyAscii = 13 Then
        txtNombreOrdena.SetFocus
    End If
End Sub

Private Sub txtNDOIRealiza_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtNDOIRealiza, KeyAscii, 30)
    If KeyAscii = 13 Then
        txtNombreRealiza.SetFocus
    End If
End Sub

Private Sub txtNombreBeneficia_Change()
    If txtNombreBeneficia.SelStart > 0 Then
        i = Len(Mid(txtNombreBeneficia.Text, 1, txtNombreBeneficia.SelStart))
    End If
    txtNombreBeneficia.Text = UCase(txtNombreBeneficia.Text)
    txtNombreBeneficia.SelStart = i
End Sub

Private Sub txtNombreBeneficia_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras(KeyAscii, True)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
End Sub

Private Sub txtNombreOrdena_Change()
    If txtNombreOrdena.SelStart > 0 Then
        i = Len(Mid(txtNombreOrdena.Text, 1, txtNombreOrdena.SelStart))
    End If
    txtNombreOrdena.Text = UCase(txtNombreOrdena.Text)
    txtNombreOrdena.SelStart = i
End Sub

Private Sub txtNombreOrdena_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras(KeyAscii, True)
If KeyAscii = 13 Then
    If Trim(Right(cmdPresencia.Text, 5)) = "1" Then
        cboTipoDOIRealiza.SetFocus
    Else
         cboTipoDOIBeneficia.SetFocus
    End If
End If
End Sub

Private Sub txtNombreRealiza_Change()
    If txtNombreRealiza.SelStart > 0 Then
        i = Len(Mid(txtNombreRealiza.Text, 1, txtNombreRealiza.SelStart))
    End If
    txtNombreRealiza.Text = UCase(txtNombreRealiza.Text)
    txtNombreRealiza.SelStart = i
End Sub

Private Sub txtNombreRealiza_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras(KeyAscii, True)
If KeyAscii = 13 Then
    cboTipoDOIBeneficia.SetFocus
End If
End Sub
