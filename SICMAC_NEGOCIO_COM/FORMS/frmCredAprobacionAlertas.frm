VERSION 5.00
Begin VB.Form frmCredAprobacionAlertas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alerta Aprobación"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "frmCredAprobacionAlertas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraVisto 
      Caption         =   "Visto de Aprobación"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   5175
      Begin VB.TextBox TxtClave 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C16A0B&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   4
         ToolTipText     =   "Ingrese su Clave Secreta"
         Top             =   840
         Width           =   2430
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Validar"
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox txtUsuario 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C16A0B&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   2
         ToolTipText     =   "Ingrese su Usuario"
         Top             =   480
         Width           =   2430
      End
      Begin VB.Label lblusuario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario :"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   495
         Width           =   630
      End
      Begin VB.Label lblClave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave     :"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   675
      End
   End
   Begin VB.Label lblMensaje 
      Caption         =   "@lblMensaje"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5115
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmCredAprobacionAlertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fbValida As Boolean
Private fbClaveEsCorrecta  As Boolean
    
Private Sub CmdAceptar_Click()
If Trim(TxtClave.Text) = "" Then
    MsgBox "Ingrese la Clave", vbInformation, "Aviso"
    TxtClave.SetFocus
Else

Dim sDominio As String
Dim oConsSist As COMDConstSistema.NCOMConstSistema
Set oConsSist = New COMDConstSistema.NCOMConstSistema
sDominio = oConsSist.LeeConstSistema(37)

fbClaveEsCorrecta = ClaveIncorrectaNT(Trim(txtUsuario.Text), TxtClave.Text, sDominio)
If Not fbClaveEsCorrecta Then
    MsgBox "Usuario y clave incorrecta", vbInformation, "Aviso"
    fbValida = False
    TxtClave.SetFocus
Else
    fbValida = True
    Unload Me
End If
Set oConsSist = Nothing
End If
End Sub

Private Sub Form_Load()
fbValida = False
fbClaveEsCorrecta = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not fbClaveEsCorrecta Then
    fbValida = False
    fbClaveEsCorrecta = False
End If
End Sub

Public Function Inicio(ByVal psMensaje As String) As Boolean
Dim ValorH As Long
Dim sValorArr() As String

sValorArr = Split(psMensaje, vbNewLine)
ValorH = UBound(sValorArr)
ValorH = 220 * (ValorH + 1)
lblMensaje.Height = ValorH
fraVisto.Top = fraVisto.Top + ValorH
Me.Height = Me.Height + ValorH
lblMensaje.Caption = psMensaje
txtUsuario.Text = gsCodUser
Me.Show 1
Inicio = fbValida
End Function

Private Sub TxtClave_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CmdAceptar_Click
End If
End Sub
