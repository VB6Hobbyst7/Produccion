VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmGeneraProtocoloOperaciones 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   Icon            =   "frmGeneraProtocoloOperaciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   3720
      Top             =   1440
   End
   Begin VB.Frame fraFecha 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   630
      Left            =   90
      TabIndex        =   3
      Top             =   30
      Width           =   3495
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   570
         TabIndex        =   4
         Top             =   285
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   1905
      TabIndex        =   1
      Top             =   1485
      Width           =   1755
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "&Procesar"
      Height          =   420
      Left            =   90
      TabIndex        =   0
      Top             =   1470
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      Height          =   840
      Left            =   105
      TabIndex        =   2
      Top             =   600
      Width           =   3480
      Begin VB.TextBox txtUsuario 
         Height          =   330
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "txtPrincipal"
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmGeneraProtocoloOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************************'
'***     Rutina:              frmGeneraProtocoloOperaciones                                  ***'
'***     Descripcion:         Permite ingresar los criterios de fecha de proceso y usuario   ***'
'***     Creado por:          PEAC - Pedro Acuña                                             ***'
'***     Maquina:             TIF-1-06                                                       ***'
'***     Fecha-Tiempo:        20100301 10:00:00 AM                                           ***'
'***     Ultima Modificacion: Creacion del Formulario                                        ***'
'***********************************************************************************************'

Option Explicit
Dim j As Long
Dim ntotal As Long
Dim lnesta As Integer, copi As Integer, I As Integer, k As Integer
Dim txt As String, intento As Integer

Private Sub cmdProcesar_Click()
Dim I As Integer
Dim Mfecha As String
Dim sCad As String
Dim oPrevio As previo.clsprevio

If txtFecha <> "__/__/____" Then
    Mfecha = ValidaFecha(txtFecha.Text)
    If Mfecha <> "" Then
        MsgBox Mfecha, vbInformation, "Aviso"
        Me.txtFecha.SetFocus
        Exit Sub
    End If
End If
If txtFecha = "__/__/____" Then
    MsgBox "Por favor Ingrese una Fecha", vbInformation, "Aviso"
    txtFecha.SetFocus
    Exit Sub
End If

Dim oProt As COMNCaptaGenerales.NCOMCaptaReportes
Set oProt = New COMNCaptaGenerales.NCOMCaptaReportes
sCad = oProt.ProtocoloOperaciones("PROTOCOLO DE USUARIO SOLES", 0, 0, gsNomAge, gcEmpresa, Me.txtFecha, gMonedaNacional, Me.txtUsuario, , Format(Me.txtFecha, gsFormatoFechaView), gsCodAge)
sCad = sCad & oProt.ProtocoloOperaciones("PROTOCOLO DE USUARIO DOLARES", 0, 0, gsNomAge, gcEmpresa, Me.txtFecha, gMonedaExtranjera, Me.txtUsuario, , Format(Me.txtFecha, gsFormatoFechaView), gsCodAge)

Set oPrevio = New previo.clsprevio
oPrevio.Show sCad, "PROTOCOLO DE USUARIO", True
Set oPrevio = Nothing

End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
If I <= 50 Then
    Me.Caption = Space(50 - I) + Left(txt, I)
    I = I + 1
Else
    If k < Len(txt) Then
        Me.Caption = Mid(txt, k)
        k = k + 1
    Else
        I = 1
        k = 1
    End If
End If
End Sub

Private Sub txtFecha_GotFocus()
    fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            Me.txtFecha.SetFocus
    End If
End Sub

Private Sub Form_Load()

Dim nModiUsu As Integer
Dim oAho As COMDCaptaGenerales.DCOMCaptaGenerales
Set oAho = New COMDCaptaGenerales.DCOMCaptaGenerales

nModiUsu = oAho.GetVisualizaSaldoPosicion(gsCodCargo)
Set oAho = Nothing
   
If nModiUsu = 1 Then
    Me.txtUsuario.Enabled = True
Else
    Me.txtUsuario.Enabled = False
End If

lnesta = 1: copi = 0: I = 1: k = 1: intento = 1
txt = "Protocolo de Operaciones"
Me.Timer1.Interval = 90

Me.txtFecha.Text = gdFecSis
Me.txtUsuario.Text = gsCodUser

End Sub

Private Sub txtUsuario_Change()
    txtUsuario.Text = UCase(txtUsuario.Text)
    I = Len(txtUsuario.Text)
    txtUsuario.SelStart = I
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub
