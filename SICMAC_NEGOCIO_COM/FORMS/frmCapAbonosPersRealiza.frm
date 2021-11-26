VERSION 5.00
Begin VB.Form frmCapAbonosPersRealiza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Persona Abono"
   ClientHeight    =   1905
   ClientLeft      =   3885
   ClientTop       =   5445
   ClientWidth     =   7380
   Icon            =   "frmCapAbonosPersRealiza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   7380
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame fraRem 
      Caption         =   "Persona que Realiza el Deposito"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7155
      Begin SICMACT.TxtBuscar txtDNI 
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Top             =   318
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
         TipoBusPers     =   1
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Doc. ID.:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   378
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   915
         Width           =   645
      End
      Begin VB.Label lblNombre 
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
         Left            =   1140
         TabIndex        =   1
         Top             =   840
         Width           =   5475
      End
   End
End
Attribute VB_Name = "frmCapAbonosPersRealiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'JACA 20110317*****************************************************************
Option Explicit
Private lsPersCod As String
Private lsPersNombre As String
Private lsPersDNI As String
Private lsPersRegistrar As Boolean
Dim op As Integer
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
Property Let PersDNI(pPersDNI As String)
   lsPersDNI = pPersDNI
End Property
Property Get PersDNI() As String
    PersDNI = lsPersDNI
End Property
Property Let PersRegistrar(pPersReg As String)
   lsPersRegistrar = pPersReg
End Property
Property Get PersRegistrar() As String
    PersRegistrar = lsPersRegistrar
End Property

Public Sub inicia()
    cmdGrabar.Enabled = False
    op = 0
    Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
    lsPersRegistrar = False
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
 lsPersRegistrar = True
 op = 1
 Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If op = 0 Then
        lsPersRegistrar = False
    End If
End Sub
Public Sub insertarPersonaDeposita(ByVal pnNroMov As Long, ByVal psPersCod As String, ByVal psPersDNI As String, ByVal psPersNombre As String)
    Dim objCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set objCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    objCap.insertarCapAbonoPersRea pnNroMov, psPersCod, psPersDNI, psPersNombre
    Set objCap = Nothing
End Sub
Private Sub txtDNI_EmiteDatos()

    If Trim(txtDNI.sPersNroDoc) = "" Then
        MsgBox "Esta persona no tiene un documento de identidad ingresado." & vbCrLf & " Por favor actualice su información.", vbOKOnly + vbInformation, "Atención"
        LimpiaControles
        Exit Sub
    End If

    If txtDNI.Text <> "" And txtDNI.sPersNroDoc <> "" Then
        
        If txtDNI.PersPersoneria = gPersonaNat Then
            lblNombre = txtDNI.psDescripcion
            txtDNI.Text = txtDNI.sPersNroDoc
            lsPersCod = txtDNI.psCodigoPersona
            lsPersNombre = txtDNI.psDescripcion
            lsPersDNI = txtDNI.sPersNroDoc
            Dim objCap As COMDCaptaGenerales.DCOMCaptaGenerales
            Set objCap = New COMDCaptaGenerales.DCOMCaptaGenerales
            If objCap.GetNroOPeradoras(gsCodAge) > 1 Then
                If txtDNI.psCodigoPersona = gsCodPersUser Then
                    MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                    LimpiaControles
                    Exit Sub
                End If
            End If
        
            Set objCap = Nothing
            cmdGrabar.Enabled = True
            cmdGrabar.SetFocus
        Else
            MsgBox "Solo se pueden Registrar Personas Naturales", vbInformation
            LimpiaControles
        End If
    End If
End Sub
Private Sub LimpiaControles()
    lblNombre = ""
    txtDNI.Text = ""
    lsPersCod = ""
    cmdGrabar.Enabled = False
End Sub
'END JACA ***********************************************************************
