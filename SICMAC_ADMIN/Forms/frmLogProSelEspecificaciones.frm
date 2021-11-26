VERSION 5.00
Begin VB.Form frmLogProSelEspecificaciones 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   2070
   ClientLeft      =   2550
   ClientTop       =   3840
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDescripcion 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   1620
      Width           =   1095
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Especificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1260
   End
   Begin VB.Shape shpMarco 
      BorderColor     =   &H00808080&
      Height          =   2055
      Left            =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmLogProSelEspecificaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vpTexto As String
Dim nX As Integer, nY As Integer, cTexto As String
Dim cTitulo As String, nAncho As Integer, nAlto As Integer

Public Sub Inicio(ByVal pnX As Integer, ByVal pnY As Integer, ByVal psTexto As String, _
               Optional psTitulo As String = "Especificación", _
               Optional pnAncho As Integer = 0, _
               Optional pnAlto As Integer = 0)
nX = pnX
nY = pnY
nAncho = pnAncho
nAlto = pnAlto
cTitulo = psTitulo
cTexto = psTexto
Me.Show 1
End Sub

Private Sub Form_Load()
txtDescripcion.Text = cTexto
lblTitulo.Caption = cTitulo
Me.vpTexto = ""
Me.Left = nX
Me.Top = nY
If nAncho > 0 Then Me.Width = nAncho
If nAlto > 0 Then Me.Height = nAlto
End Sub

Private Sub cmdCerrar_Click()
Me.vpTexto = txtDescripcion.Text
Unload Me
End Sub

Private Sub Form_Resize()
txtDescripcion.Move 120, 360, Me.Width - 360, Me.Height - 1260
shpMarco.Move 0, 0, Me.Width, Me.Height
End Sub

Private Sub txtDescripcion_GotFocus()
SelTexto txtDescripcion
End Sub
