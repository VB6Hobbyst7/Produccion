VERSION 5.00
Begin VB.Form frmRHBuscarEmpleado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Empleado"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "frmRHBuscarEmpleado.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   345
      Left            =   1560
      TabIndex        =   1
      Top             =   405
      Width           =   1110
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   2760
      TabIndex        =   2
      Top             =   405
      Width           =   1110
   End
   Begin VB.TextBox txtNombre 
      Height          =   300
      Left            =   60
      MaxLength       =   50
      TabIndex        =   0
      Top             =   30
      Width           =   5265
   End
End
Attribute VB_Name = "frmRHBuscarEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsNombre As String

Private Sub CmdBuscar_Click()
    lsNombre = Me.txtNombre.Text
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    lsNombre = ""
    Unload Me
End Sub

Private Sub txtNombre_GotFocus()
    txtNombre.SelStart = 0
    txtNombre.SelLength = 50
End Sub

Public Function GetNombre(psBuscar As String) As String
    txtNombre.Text = psBuscar
    Me.Show 1
    GetNombre = lsNombre
End Function

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub
