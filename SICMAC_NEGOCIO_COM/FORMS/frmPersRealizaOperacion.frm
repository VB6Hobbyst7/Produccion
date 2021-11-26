VERSION 5.00
Begin VB.Form frmPersRealizaOperacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orden de Pago – Registro de Usuario"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   Icon            =   "frmPersRealizaOperacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "&Limpiar"
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
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
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
      Left            =   6240
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Left            =   5040
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame fraDatosPersRealiza 
      Caption         =   "Persona Realiza la Operación"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtOrigen 
         Height          =   690
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   5895
      End
      Begin VB.CommandButton cmdBuscarPers 
         Caption         =   "..."
         Height          =   330
         Left            =   6720
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtNDOI 
         Height          =   330
         Left            =   1080
         TabIndex        =   3
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ComboBox cboTipoDOI 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         Height          =   330
         Left            =   1080
         MaxLength       =   200
         TabIndex        =   0
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label lblOrigen2 
         AutoSize        =   -1  'True
         Caption         =   "Fondos:"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   2160
         Width           =   570
      End
      Begin VB.Label lblOrigen 
         AutoSize        =   -1  'True
         Caption         =   "Origen de"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº DOI:"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   1400
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo DOI:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   915
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   380
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmPersRealizaOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lsPersCod As String
Private lsPersNombre As String
Private lsPersDOI As String
Private lsPersTipoDOI As String
Private lsTipoCliente As String
Private lsPersRegistrar As Boolean
Private lsTipoOperacion As Integer
Private lsOrigenFondos As String
Dim fbPersNatural As Boolean
Dim i As Integer
Property Let PersCod(pPersCod As String)
   lsPersCod = pPersCod
End Property
Property Get PersCod() As String
    PersCod = lsPersCod
End Property
Property Let PersNombre(pPersNombre As String)
   lsPersNombre = pPersNombre
End Property
Property Get PersNombre() As String
    PersNombre = lsPersNombre
End Property
Property Let PersDOI(pPersDOI As String)
   lsPersDOI = pPersDOI
End Property
Property Get PersDOI() As String
    PersDOI = lsPersDOI
End Property
Property Let PersTipoDOI(pPersTipoDOI As String)
   lsPersTipoDOI = pPersTipoDOI
End Property
Property Get PersTipoDOI() As String
    PersTipoDOI = lsPersTipoDOI
End Property
Property Let PersTipoCliente(pPersTipoCliente As String)
   lsTipoCliente = pPersTipoCliente
End Property
Property Get PersTipoCliente() As String
    PersTipoCliente = lsTipoCliente
End Property
Property Let PersRegistrar(pPersReg As String)
   lsPersRegistrar = pPersReg
End Property
Property Get PersRegistrar() As String
    PersRegistrar = lsPersRegistrar
End Property
Property Let TipoOperacion(pTpoOpe As Integer)
   lsTipoOperacion = pTpoOpe
End Property
Property Get TipoOperacion() As Integer
    TipoOperacion = lsTipoOperacion
End Property
'WIOR 20121114 ************
Property Let Origen(psOrigen As String)
   lsOrigenFondos = psOrigen
End Property
Property Get Origen() As String
    Origen = lsOrigenFondos
End Property
'WIOR FIN *****************
Private Sub cboTipoDOI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNDOI.SetFocus
End If
End Sub

Private Sub cmdBuscarPers_Click()
Dim oPers As COMDPersona.UCOMPersona
Dim oPersona As UPersona_Cli
Dim nTipoDOI As Integer
Dim sTipoDOI As String
Dim sNumeroDOI As String

Call HabilitaControles(True, 1)

    
    Set oPers = frmBuscaPersona.Inicio
    Set oPersona = New UPersona_Cli
    If Not oPers Is Nothing Then
        oPersona.RecuperaPersona (oPers.sPersCod)
    Else
        Call HabilitaControles(True, 1)
        Exit Sub
    End If
    
    If Not oPersona Is Nothing Then
    
        If oPersona.Personeria = "1" Then
            Call oPersona.ObtenerDatosDocumentoxPos(0, nTipoDOI, sTipoDOI, sNumeroDOI)
            txtNombre.Text = PstaNombre(oPersona.NombreCompleto, True)
            cboTipoDOI.ListIndex = IndiceListaCombo(cboTipoDOI, nTipoDOI)
            txtNDOI.Text = sNumeroDOI
            Call HabilitaControles(False)
            lsPersCod = oPers.sPersCod
            lsPersNombre = PstaNombre(oPersona.NombreCompleto, True)
            lsPersDOI = sNumeroDOI
            lsPersTipoDOI = nTipoDOI
            lsTipoCliente = 1
            cmdGrabar.SetFocus
        Else
            MsgBox "No es una Persona Natural", vbInformation, "Aviso"
            Call HabilitaControles(True, 1)
            lsTipoCliente = 0
            Exit Sub
        End If

    Else
        Call HabilitaControles(True, 1)
        Exit Sub
    End If
    Set oPersona = Nothing
End Sub

Private Sub cmdCancelar_Click()
    lsPersRegistrar = False
    Unload Me
End Sub

Private Sub CmdGrabar_Click()
If ValidaDatos Then
    If MsgBox("Esta seguro de grabar los datos ", vbInformation + vbYesNo, "Aviso") = vbYes Then
    If lsTipoCliente = 0 Then
        lsPersCod = ""
        lsPersNombre = Trim(txtNombre.Text)
        lsPersDOI = Trim(txtNDOI.Text)
        lsPersTipoDOI = Trim(Right(cboTipoDOI.Text, 5))
    End If
    'WIOR 20121114 ***************************
    If Me.txtOrigen.Visible Then
        lsOrigenFondos = Trim(txtOrigen.Text)
    Else
        lsOrigenFondos = ""
    End If
    'WIOR FIN ********************************
    lsPersRegistrar = True
    Unload Me
    End If
End If
End Sub

Private Sub CargaControles()
    Dim oConstante As COMDConstantes.DCOMConstantes
    Set oConstante = New COMDConstantes.DCOMConstantes
    Call Llenar_Combo_con_Recordset(oConstante.RecuperaConstantes(gPersIdTipo), cboTipoDOI)
    lsTipoCliente = 0
    lsPersRegistrar = False
End Sub

Private Sub CmdLimpiar_Click()
Call HabilitaControles(True, 1)
End Sub
Public Sub Inicia(ByVal psDesc As String, ByVal pnTipo As TipoOperacionPersRealiza, Optional ByVal pnCliReforzado As Integer = 0) 'WIOR 20121114 AGREGO pnCliReforzado
Me.Caption = psDesc & " – Registro de Usuario"
lsTipoOperacion = pnTipo
Call CargaControles
'WIOR 20121114*******************
If pnCliReforzado = 0 Then
    Me.Height = 3200
    fraDatosPersRealiza.Height = 1900
    lblOrigen.Visible = False
    txtOrigen.Visible = False
    cmdLimpiar.Top = 2200
    cmdGrabar.Top = 2200
    cmdCancelar.Top = 2200
ElseIf pnCliReforzado = 1 Then
    Me.Height = 4080
    fraDatosPersRealiza.Height = 2775
    lblOrigen.Visible = True
    txtOrigen.Visible = True
    cmdLimpiar.Top = 3120
    cmdGrabar.Top = 3120
    cmdCancelar.Top = 3120
End If
'WIOR FIN ***********************
Me.Show 1
End Sub
Private Sub HabilitaControles(ByVal pbHabilita As Boolean, Optional ByVal pnLimpia As Integer = 0)
txtNombre.Enabled = pbHabilita
txtNDOI.Enabled = pbHabilita
cboTipoDOI.Enabled = pbHabilita
Me.txtOrigen.Text = "" 'WIOR 20121114*
If pnLimpia <> 0 Then
    txtNombre.Text = ""
    txtNDOI.Text = ""
    Call CargaControles
    cboTipoDOI.ListIndex = IndiceListaCombo(cboTipoDOI, 0)
    lsTipoCliente = 0
    txtNombre.SetFocus
End If

End Sub
Private Function ValidaDatos() As Boolean
If lsTipoCliente = 0 Then
Dim ClsPersona As COMDPersona.DCOMPersonas
Set ClsPersona = New COMDPersona.DCOMPersonas
Dim R As ADODB.Recordset

    If Trim(txtNombre.Text) = "" Then
        MsgBox "Ingrese el Nombre de la Persona que esta realizando el retiro.", vbInformation, "Aviso"
        txtNombre.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    If Trim(Me.cboTipoDOI.Text) = "" Then
        MsgBox "Seleccione el Tipo de DOI de la Persona que esta realizando el retiro.", vbInformation, "Aviso"
        cboTipoDOI.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    If Trim(txtNDOI.Text) = "" Then
        MsgBox "Ingrese el Nro. de DOI de la Persona que esta realizando el retiro.", vbInformation, "Aviso"
        cboTipoDOI.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    
    'Verfica Longitud de DOI
    If Trim(cboTipoDOI.Text) <> "" Then
        If CInt(Trim(Right(cboTipoDOI.Text, 5))) = gPersIdDNI Then
            If Len(Trim(txtNDOI.Text)) <> gnNroDigitosDNI Then
                MsgBox "DNI No es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
                txtNDOI.SetFocus
                ValidaDatos = False
                Exit Function
            End If
        End If
        If CInt(Trim(Right(cboTipoDOI.Text, 5))) = gPersIdRUC Then
            If Len(Trim(txtNDOI.Text)) <> gnNroDigitosRUC Then
                MsgBox "RUC No es de " & gnNroDigitosRUC & " digitos", vbInformation, "Aviso"
                txtNDOI.SetFocus
                ValidaDatos = False
                Exit Function
            End If
        End If
    End If
   
    Set R = ClsPersona.BuscaCliente(txtNDOI.Text, 3)
    If R.RecordCount > 0 Then
        MsgBox "Nro. de DOI " & txtNDOI.Text & " ya existe el la base de datos, Favor de Buscar a la persona como cliente CMACMaynas.", vbInformation, "Aviso"
        txtNDOI.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    Set R = Nothing
End If
'WIOR 20121114 **********************************************************************************
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
'WIOR FIN ****************************************************************************************

ValidaDatos = True
End Function

Private Sub txtNDOI_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtNDOI, KeyAscii, 30)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub

Private Sub txtNombre_Change()
If txtNombre.SelStart > 0 Then
    i = Len(Mid(txtNombre.Text, 1, txtNombre.SelStart))
End If
txtNombre.Text = UCase(txtNombre.Text)
txtNombre.SelStart = i
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cboTipoDOI.SetFocus
End If
End Sub
Public Sub InsertaPersonaRealizaOperacion(ByVal pnNroMov As Long, ByVal pnCtaCod As String, ByVal pnTipoCliente As Integer, _
                                            ByVal psPersCod As String, ByVal pnPersDOITipo As Integer, ByVal pnPersDOI As String, _
                                            ByVal psPerNombre As String, ByVal pnTipo As Integer, _
                                            Optional ByVal pcOrigen As String = "", Optional ByVal pnCondicion As Integer = 0) 'WIOR 20121114 AGREGO pcOrigen,pnCondicion
    Dim oNPersona As COMNPersona.NCOMPersona
    Set oNPersona = New COMNPersona.NCOMPersona
    Call oNPersona.InsertaPersonaRealizaOperacion(pnNroMov, pnCtaCod, pnTipoCliente, psPersCod, pnPersDOITipo, pnPersDOI, psPerNombre, pnTipo, pcOrigen, pnCondicion) 'WIOR 20121114 AGREGO pcOrigen,pnCondicion
    Set oNPersona = Nothing
End Sub
