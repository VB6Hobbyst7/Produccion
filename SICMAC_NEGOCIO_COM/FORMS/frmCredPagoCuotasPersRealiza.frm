VERSION 5.00
Begin VB.Form frmCredPagoCuotasPersRealiza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Persona Pago Cuotas"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
   Icon            =   "frmCredPagoCuotasPersRealiza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Persona que realiza el Pago de la Cuota"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin SICMACT.TxtBuscar txtDNI 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.Label lblNombre 
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
         Height          =   345
         Left            =   1200
         TabIndex        =   4
         Top             =   885
         Width           =   5475
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Doc. ID.:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   420
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmCredPagoCuotasPersRealiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**Juez 20120323 ************************************
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

Public Sub Inicia()
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
Public Sub insertaPersonaRealizaPagoCuotas(ByVal pnNroMov As Long, ByVal psPersCod As String, ByVal psPersDNI As String, ByVal psPersNombre As String)
    Dim oCred As COMNCredito.NCOMCredito
    Set oCred = New COMNCredito.NCOMCredito
    oCred.insertaPersonaRealizaPagoCuotas pnNroMov, psPersCod, psPersDNI, psPersNombre
    Set oCred = Nothing
End Sub
Private Sub txtDNI_EmiteDatos()

    If Trim(txtDNI.sPersNroDoc) = "" Then
        MsgBox "No seleccionó a la Persona.", vbOKOnly + vbInformation, "Atención"
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
            'Dim objCap As COMDCaptaGenerales.DCOMCaptaGenerales
            'Set objCap = New COMDCaptaGenerales.DCOMCaptaGenerales
            'If objCap.GetNroOPeradoras(gsCodAge) > 1 Then
            If txtDNI.psCodigoPersona = gsCodPersUser Then
                MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                LimpiaControles
                Exit Sub
            End If
            'End If
        
            'Set objCap = Nothing
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
'**End Juez ******************************************
