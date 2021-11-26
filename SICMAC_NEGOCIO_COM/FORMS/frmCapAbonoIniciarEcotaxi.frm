VERSION 5.00
Begin VB.Form frmCapAbonoIniciarEcotaxi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de cliente que realiza el abono"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   FillStyle       =   0  'Solid
   Icon            =   "frmCapAbonoIniciarEcotaxi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Iniciar Ecotaxi"
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
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtDocumento 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtNombre 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   5175
      End
      Begin SICMACT.TxtBuscar txtBuscaPersona 
         Height          =   300
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   529
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
      End
      Begin VB.Label Label2 
         Caption         =   "Documento :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre  :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Codigo   :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCapAbonoIniciarEcotaxi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCodigoPersona As String
Dim lsNombrePersona As String
Dim lsPersDNIEcotaxi As String
Public Sub Inicio(ByRef psCodigoPersona As String, ByRef psNombrePersona As String, ByRef psPersDNIEcotaxi As String)
    lsCodigoPersona = psCodigoPersona
    lsNombrePersona = psNombrePersona
    lsPersDNIEcotaxi = psPersDNIEcotaxi
    Call CentraForm(Me)
    Show 1
    psCodigoPersona = lsCodigoPersona
    psNombrePersona = lsNombrePersona
    psPersDNIEcotaxi = lsPersDNIEcotaxi
End Sub

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    txtNombre.Text = ""
    txtDocumento.Text = ""
    txtBuscaPersona.Text = ""
    lsCodigoPersona = ""
    lsNombrePersona = ""
    lsPersDNIEcotaxi = ""
    MsgBox "Para proceder al pago de la inicial del credito ecotaxi debe selecionarse " & _
            " el cliente que realiza el abono, si no es un pago inicial ecotaxi, deseleccione " & _
            " el checkbox Inicial ecotaxi"
    Unload Me
End Sub

Private Sub txtBuscaPersona_EmiteDatos()
    Dim ClsPersona As COMDPersona.DCOMPersonas
    Dim Rf As ADODB.Recordset
    Set Rf = New ADODB.Recordset
    Set ClsPersona = New COMDPersona.DCOMPersonas
    If Trim(txtBuscaPersona.Text) <> "" Then
    Set Rf = ClsPersona.BuscaCliente(txtBuscaPersona.Text, BusquedaCodigo)
    txtNombre.Text = Rf!cPersNombre
    txtDocumento.Text = IIf(Rf!nPersPersoneria = 1, Rf!cPersIDnroDNI, Rf!cPersIDnroRUC)
    lsCodigoPersona = txtBuscaPersona.Text
    lsNombrePersona = txtNombre.Text
    lsPersDNIEcotaxi = txtDocumento.Text
    End If
End Sub
