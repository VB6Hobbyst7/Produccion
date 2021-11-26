VERSION 5.00
Begin VB.Form frmDescriOpeFrecuente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descripcion de Operaciones Frecuentes"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6465
   Icon            =   "frmDescriOpeFrecuente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtDescrip 
      Height          =   325
      Left            =   1200
      TabIndex        =   0
      Top             =   145
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Descipcion :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmDescriOpeFrecuente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lsDescrip As String

Private Sub cmdAceptar_Click()
    lsDescrip = txtDescrip.Text
    If lsDescrip = "" Then
        MsgBox "La descipcion no puede ser en blanco", vbInformation, "MENSAJE DEL SISTEMA"
        txtDescrip.SetFocus
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        lsDescrip = ""
    End If
End Sub
Private Sub txtDescrip_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub
