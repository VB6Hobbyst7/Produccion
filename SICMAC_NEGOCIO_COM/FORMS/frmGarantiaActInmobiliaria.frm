VERSION 5.00
Begin VB.Form frmGarantiaActInmobiliaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Datos de la Garantía"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   Icon            =   "frmGarantiaActInmobiliaria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3945
      TabIndex        =   13
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   1920
      Width           =   1000
   End
   Begin VB.Frame fraDescripcion 
      Caption         =   "Descripción"
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
      Height          =   1815
      Left            =   75
      TabIndex        =   0
      Top             =   40
      Width           =   4845
      Begin VB.TextBox txtAnioConstruccion 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   14
         Tag             =   "txtPrincipal"
         Text            =   "2014"
         Top             =   1440
         Width           =   650
      End
      Begin VB.TextBox txtNroSotanos 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   10
         Tag             =   "txtPrincipal"
         Text            =   "0"
         Top             =   1440
         Width           =   650
      End
      Begin VB.ComboBox cmbInmuebleCate 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1080
         Width           =   3040
      End
      Begin VB.TextBox txtNroPisos 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "txtPrincipal"
         Text            =   "0"
         Top             =   720
         Width           =   650
      End
      Begin VB.TextBox txtNroLocales 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "txtPrincipal"
         Text            =   "0"
         Top             =   720
         Width           =   650
      End
      Begin VB.ComboBox cmbInmuebleClase 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   3040
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Año de Construcción :"
         Height          =   195
         Left            =   2480
         TabIndex        =   11
         Top             =   1470
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N° de Locales :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   750
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Categoría :"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1100
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N° de Sótanos :"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1470
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° de Pisos :"
         Height          =   195
         Left            =   2760
         TabIndex        =   2
         Top             =   750
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Clase de Inmueble :"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmGarantiaActInmobiliaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
'** Nombre : frmGarantiaActInmobiliaria
'** Descripción : Para actualización de Garantía Inmobiliaria segun TI-ERS063-2014
'** Creación : EJVG, 20150605 12:00:00 PM
'*********************************************************************************
Option Explicit

Dim fsNumGarant As String
Dim fdValorizacion As Date

Dim fbAceptar As Boolean

Public Function Actualizar(ByVal psNumGarant As String, ByVal pdValorizacion As Date) As Boolean
    fsNumGarant = psNumGarant
    fdValorizacion = pdValorizacion
    Show 1
    Actualizar = fbAceptar
End Function
Private Sub cmbInmuebleCate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtNroSotanos
    End If
End Sub
Private Sub cmbInmuebleClase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtNroLocales
    End If
End Sub

Private Sub Form_Load()
    Dim oGarantia As New COMDCredito.DCOMGarantia
    Dim rs As New ADODB.Recordset
    
    Screen.MousePointer = 11
    
    CargarControles
    Limpiar
    
    Set rs = oGarantia.RecuperaInmueblexSolicitudPoliza(fsNumGarant, fdValorizacion)
    If Not rs.EOF Then
        cmbInmuebleClase.ListIndex = IndiceListaCombo(cmbInmuebleClase, rs!nClase)
        txtNroLocales.Text = rs!nNroLocales
        txtNroPisos.Text = rs!nNroPisos
        cmbInmuebleCate.ListIndex = IndiceListaCombo(cmbInmuebleClase, rs!nCategoria)
        txtNroSotanos.Text = rs!nNroSotanos
        txtAnioConstruccion.Text = rs!nAnioConstruccion
    Else
        MsgBox "No se ha encontrado datos del Inmueble", vbExclamation, "Aviso"
        cmdAceptar.Enabled = False
    End If
    RSClose rs
    Screen.MousePointer = 0
End Sub
Private Sub cmdCancelar_Click()
    fbAceptar = False
    Unload Me
End Sub
Private Sub CargarControles()
    Dim oCons As New COMDConstantes.DCOMConstantes
    Dim rs As New ADODB.Recordset
       
    Set rs = oCons.RecuperaConstantes(gGarantiaClaseTasacInmobiliaria)
    Call Llenar_Combo_con_Recordset(rs, cmbInmuebleClase)
    
    Set rs = oCons.RecuperaConstantes(gGarantiaCategoriaTasacInmobiliaria)
    Call Llenar_Combo_con_Recordset(rs, cmbInmuebleCate)
        
    RSClose rs
    Set oCons = Nothing
End Sub
Private Sub Limpiar()
    cmbInmuebleClase.ListIndex = -1
    txtNroLocales.Text = "0"
    txtNroPisos.Text = "0"
    cmbInmuebleCate.ListIndex = -1
    txtNroSotanos.Text = "0"
    txtAnioConstruccion.Text = Year(gdFecSis)
End Sub
Private Sub txtNroLocales_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl txtNroPisos
    End If
End Sub
Private Sub txtNroPisos_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl cmbInmuebleCate
    End If
End Sub
Private Sub txtNroSotanos_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl txtAnioConstruccion
    End If
End Sub
Private Sub txtAnioConstruccion_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl cmdAceptar
    End If
End Sub
Private Function validarDatos() As Boolean
    If cmbInmuebleClase.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la Clase del Inmueble", vbInformation, "Aviso"
        EnfocaControl cmbInmuebleClase
        Exit Function
    End If
    If cmbInmuebleCate.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la Categoría del Inmueble", vbInformation, "Aviso"
        EnfocaControl cmbInmuebleCate
        Exit Function
    End If
    If Not IsNumeric(txtNroLocales.Text) Then
        MsgBox "Ud. debe de especificar el N° de Locales", vbInformation, "Aviso"
        EnfocaControl txtNroLocales
        Exit Function
    Else
        If CInt(txtNroLocales.Text) <= 0 Then
            MsgBox "El N° de Locales debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtNroLocales
            Exit Function
        End If
    End If
    If Not IsNumeric(txtNroPisos.Text) Then
        MsgBox "Ud. debe de especificar el N° de Pisos", vbInformation, "Aviso"
        EnfocaControl txtNroPisos
        Exit Function
    Else
        If CInt(txtNroPisos.Text) <= 0 Then
            MsgBox "El N° de Pisos debe ser mayor a cero", vbInformation, "Aviso"
            EnfocaControl txtNroPisos
            Exit Function
        End If
    End If
    If Not IsNumeric(txtNroSotanos.Text) Then 'Nro de Sotanos puede ser cero
        MsgBox "Ud. debe de especificar el N° de Sotanos", vbInformation, "Aviso"
        EnfocaControl txtNroSotanos
        Exit Function
    End If
    If Not IsNumeric(txtAnioConstruccion.Text) Then
        MsgBox "Ud. debe de especificar el Año de Construcción", vbInformation, "Aviso"
        EnfocaControl txtAnioConstruccion
        Exit Function
    Else
        If CInt(txtAnioConstruccion.Text) <= 1890 Then
            MsgBox "Ud. debe de especificar el Año de Construcción", vbInformation, "Aviso"
            EnfocaControl txtAnioConstruccion
            Exit Function
        End If
        If CInt(txtAnioConstruccion.Text) > Year(gdFecSis) Then
            MsgBox "El Año de Construcción no debe ser mayor al año del sistema", vbInformation, "Aviso"
            EnfocaControl txtAnioConstruccion
            Exit Function
        End If
    End If
        
    validarDatos = True
End Function
Private Sub cmdAceptar_Click()
    Dim oGarantia As New COMNCredito.NCOMGarantia
    Dim bExito As Boolean
    
    On Error GoTo ErrAceptar
    
    cmdAceptar.Enabled = False
    If Not validarDatos Then
        cmdAceptar.Enabled = True
        Exit Sub
    End If
    
    If MsgBox("¿Está seguro de actualizar los datos de la Garantía Inmobiliaria?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        cmdAceptar.Enabled = True
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Set oGarantia = New COMNCredito.NCOMGarantia
    bExito = oGarantia.ActualizarGarantiaInmobiliaria(fsNumGarant, fdValorizacion, CInt(Trim(Right(cmbInmuebleClase.Text, 3))), CInt(txtNroLocales.Text), CInt(txtNroPisos.Text), CInt(Trim(Right(cmbInmuebleCate.Text, 3))), CInt(txtNroSotanos.Text), CInt(Me.txtAnioConstruccion.Text))
    Screen.MousePointer = 0
    cmdAceptar.Enabled = True
    Set oGarantia = Nothing
    
    If bExito Then
        MsgBox "Se ha actualizado satisfactoriamente la Garantía Inmobiliaria", vbInformation, "Aviso"
    Else
        MsgBox "Ha sucedido un error al actualizar la Garantía Inmobiliaria, si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
        Exit Sub
    End If
    
    fbAceptar = True
    Unload Me
    Exit Sub
ErrAceptar:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
