VERSION 5.00
Begin VB.Form frmGiroMantenimiento 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "frmGiroMantenimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   225
      TabIndex        =   4
      Top             =   3945
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   3945
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   3945
      Width           =   1095
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Datos del Destinatario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1920
      Left            =   105
      TabIndex        =   15
      Top             =   1920
      Width           =   7635
      Begin VB.TextBox txtnumdoc 
         Height          =   375
         Left            =   1080
         TabIndex        =   23
         Top             =   345
         Width           =   1695
      End
      Begin VB.TextBox txtNombre 
         Height          =   375
         Left            =   1080
         TabIndex        =   21
         Top             =   720
         Width           =   6435
      End
      Begin VB.TextBox txtReferencia 
         Height          =   375
         Left            =   1080
         TabIndex        =   20
         Top             =   1080
         Width           =   6435
      End
      Begin VB.TextBox txtGlosa 
         Height          =   375
         Left            =   1080
         TabIndex        =   16
         Top             =   1440
         Width           =   6435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DNI :"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Glosa:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Referencia:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   645
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Datos Giro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1710
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   7635
      Begin VB.CommandButton cmdExaminar 
         Caption         =   "&Examinar"
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   300
         Width           =   975
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   767
         Texto           =   "Giro N°"
      End
      Begin VB.Label lblSimbolo 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6885
         TabIndex        =   14
         Top             =   1215
         Width           =   345
      End
      Begin VB.Label lblAgencia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   4740
         TabIndex        =   13
         Top             =   840
         Width           =   2595
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ag. Destino :"
         Height          =   195
         Left            =   3600
         TabIndex        =   12
         Top             =   915
         Width           =   915
      End
      Begin VB.Label lblMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   4740
         TabIndex        =   11
         Top             =   1200
         Width           =   1755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   3600
         TabIndex        =   10
         Top             =   1275
         Width           =   540
      End
      Begin VB.Label lblTipo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1500
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Giro :"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1275
         Width           =   735
      End
      Begin VB.Label lblApertura 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   350
         Left            =   1500
         TabIndex        =   7
         Top             =   837
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Apertura :"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   915
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmGiroMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sOperacion As String, sRemitente As String

Dim Gtitular As String
Dim GAutNivel As String
Dim GAutMontoFinSol As Double
Dim GAutMontoFinDol As Double
Dim GMontoAprobado As Double, GNroID As Long, GPersCod As String, GClaveGiro As String
Dim lcOpeCod As String

Private Function IniciaLavDinero(poLavDinero As SICMACT.frmMovLavDinero) As String
    Dim i As Long
    Dim nRelacion As COMDConstantes.CaptacRelacPersona
    Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
    Dim nMonto As Double
    Dim sCuenta As String

    poLavDinero.TitPersLavDinero = sRemitente

    nMonto = CDbl(lblMonto)
    sCuenta = txtCuenta.NroCuenta

End Function

Private Sub ClearScreen()

    lblMonto = "0.00"
    lblTipo = ""
    lblApertura = ""
    lblAgencia = ""
    
    Me.txtnumdoc.Text = ""
    Me.txtNombre.Text = ""
    Me.txtReferencia.Text = ""
    Me.txtGlosa.Text = ""
    
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False

    fraDatos.Enabled = True
    fraCliente.Enabled = False
    txtCuenta.Cuenta = ""
    txtCuenta.Age = ""
    txtCuenta.Prod = gGiro
    txtCuenta.CMAC = gsCodCMAC
    txtCuenta.EnabledProd = False
    txtCuenta.EnabledAge = True
    txtCuenta.EnabledCta = True
    lblMonto.BackColor = &HFFFFFF
    lblSimbolo.Caption = ""
    sRemitente = ""
End Sub

Private Sub CargaDatosGiro(ByVal sCuenta As String)
    Dim rsGiro As ADODB.Recordset
    Dim clsGiro As COMNCaptaServicios.NCOMCaptaServicios

    Dim nFila As Long
    Dim sDestinatario As String
    Set clsGiro = New COMNCaptaServicios.NCOMCaptaServicios
    Set rsGiro = clsGiro.GetGiroDatos(sCuenta)
    If Not (rsGiro.EOF And rsGiro.BOF) Then
    
        If gsCodAge <> Mid(sCuenta, 4, 2) Then
            MsgBox "Solo se puede modificar los datos del Destinatario en la Agencia Origen del Giro.", vbInformation, "Aviso"
            Exit Sub
        End If
    
        If rsGiro("cNumDocDesti") = "" Then
            MsgBox "El Destinatario está registrado en el Sistema, por lo tanto debe utilizar el Módulo Persona para cambiar sus datos.", vbInformation, "Aviso"
            Exit Sub
        Else
            Me.txtnumdoc.Text = Trim(rsGiro("cNumDocDesti"))
            Me.txtNombre.Text = PstaNombre(rsGiro("cDestinatario"), False)
            Me.txtReferencia.Text = Trim(rsGiro("cReferencia"))
            Me.txtGlosa.Text = Trim(rsGiro("cGlosaMant"))
        End If
    
        sDestinatario = ""
        lblAgencia = Trim(rsGiro("cAgencia"))
        lblMonto = Format$(rsGiro("nSaldo"), "#,##0.00")
        lblTipo = Trim(rsGiro("cTipo"))
        lblApertura = Format$(rsGiro("dPrdEstado"), "dd mmm yyyy")
        sRemitente = Trim(rsGiro("cPersCod"))
    
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
        
        fraDatos.Enabled = False
        fraCliente.Enabled = True
        cmdGrabar.Enabled = True
        cmdCancelar.Enabled = True

        GClaveGiro = clsGiro.GetGiroSeguridad(sCuenta)
        If GClaveGiro <> "" Then

        End If
        cmdGrabar.SetFocus
    Else
        MsgBox "Número de Giro no encontrado o Cancelado.", vbInformation, "SICMACM - Aviso"
        txtCuenta.Age = ""
        txtCuenta.Cuenta = ""
        txtCuenta.SetFocusAge
        sRemitente = ""
    End If
    Set clsGiro = Nothing
End Sub

Private Sub cmdCancelar_Click()
    ClearScreen
    txtCuenta.SetFocus
End Sub

Private Sub cmdexaminar_Click()
frmGiroDetalleMantenimiento.Show 1
Dim sCuenta As String
Dim nmoneda As Moneda

sCuenta = txtCuenta.NroCuenta
If Len(sCuenta) = 18 Then
    txtCuenta.SetFocusCuenta
    nmoneda = CLng(Mid(sCuenta, 9, 1))
    If nmoneda = COMDConstantes.gMonedaExtranjera Then
        lblMonto.BackColor = &HC0FFC0
        lblSimbolo.Caption = "US$"
    Else
        lblMonto.BackColor = &HFFFFFF
        lblSimbolo.Caption = "S/."
    End If
    SendKeys "{Enter}"
End If
End Sub

Private Sub cmdGrabar_Click()
    Dim sCuenta As String, sMovNro As String, sPersLavDinero As String
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    
    Dim clsMantGiro As COMDCaptaServicios.DCOMCaptaServicios
    Dim bTrans As Boolean
    Dim nMovNro As Long

    If Trim(Me.txtnumdoc.Text) = "" Or Trim(Me.txtNombre.Text) = "" Or Trim(Me.txtReferencia.Text) = "" Or Trim(Me.txtGlosa.Text) = "" Then
        MsgBox "Debe llenar todos los campos.", vbInformation, "Aviso"
        Exit Sub
    End If

    If MsgBox("¿Desea Grabar los Cambios?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If

    sCuenta = txtCuenta.NroCuenta
    
    Set clsMantGiro = New COMDCaptaServicios.DCOMCaptaServicios
    
    clsMantGiro.dbCmact.BeginTrans
    bTrans = True

        Set clsMov = New COMNContabilidad.NCOMContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set clsMov = Nothing

        clsMantGiro.AgregaMov sMovNro, lcOpeCod, Trim(Me.txtGlosa.Text), gMovEstContabNoContable, gMovFlagModificado

        nMovNro = clsMantGiro.GetnMovNro(sMovNro)

        clsMantGiro.MantenimientoGiroDestinatario sCuenta, Trim(Me.txtNombre.Text), Trim(Me.txtReferencia.Text), Trim(Me.txtnumdoc.Text), nMovNro
    
    clsMantGiro.dbCmact.CommitTrans
    bTrans = False

    MsgBox "Los Cambios se Guardaron Satisfactoriamente.", vbInformation, "Aviso"

    ClearScreen
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
    frmMantCreditos.Show 1
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Me.Caption = "Mantenimiento de Giro"
    sOperacion = "Mantenimiento de Giro"
    lcOpeCod = "310400"
    ClearScreen
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCuenta As String, sMoneda As String
        sCuenta = txtCuenta.NroCuenta
        CargaDatosGiro sCuenta
    End If
End Sub

Private Sub TxtCuenta_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sCuenta As String, sMoneda As String
sCuenta = txtCuenta.NroCuenta
sMoneda = Mid(sCuenta, 9, 1)
If sMoneda = "2" Then
    lblMonto.BackColor = &HC0FFC0
    lblSimbolo.Caption = "US$"
Else
    lblMonto.BackColor = &HFFFFFF
    lblSimbolo.Caption = "S/."
End If
End Sub

Private Function Cargousu(ByVal NomUser As String) As String
 Dim rs As New ADODB.Recordset
 Dim oCons As COMDConstSistema.DCOMUAcceso
 Set oCons = New COMDConstSistema.DCOMUAcceso
 
 Set rs = oCons.Cargousu(NomUser)
  If Not (rs.EOF And rs.BOF) Then
    Cargousu = rs(0)
  End If
 Set rs = Nothing

 Set oCons = Nothing
End Function

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtReferencia.SetFocus
    End If
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        txtNombre.SetFocus
    End If
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtGlosa.SetFocus
    End If
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub
