VERSION 5.00
Begin VB.Form frmBuscaPlacaINFOGAS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFOGAS"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4260
   Icon            =   "frmBuscaPlacaINFOGAS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPlaca 
      Height          =   300
      Left            =   120
      MaxLength       =   100
      TabIndex        =   1
      Top             =   1080
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese el Número de Placa"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmBuscaPlacaINFOGAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fsPlaca As String

Public Function ObtenerPlaca(ByVal psPlaca As String) As String
    txtPlaca.Text = psPlaca
    CentraForm Me
    Me.Show 1
    ObtenerPlaca = fsPlaca
End Function
Private Sub cmdAceptar_Click()
    fsPlaca = Trim(txtPlaca.Text)
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    fsPlaca = ""
    Unload Me
End Sub
Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub
