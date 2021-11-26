VERSION 5.00
Begin VB.Form frmChequeComentarioGral 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4035
   Icon            =   "frmChequeComentarioGral.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnAccion 
      Caption         =   "&Accion"
      Height          =   340
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   340
      Left            =   2790
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtComentario 
      Height          =   960
      Left            =   80
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   80
      Width           =   3930
   End
End
Attribute VB_Name = "frmChequeComentarioGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'** Nombre : frmChequeComentarioGral
'** Descripción : Clase de Cheques  creado según RFC117-2012
'** Creación : EJVG 20121204 09:00:00 AM
'***********************************************************************
Option Explicit
Dim fsComentario As String

Public Function Inicio(ByVal pnInicio As Integer) As String
    Inicio = ""
    fsComentario = ""
    Select Case pnInicio
        Case 1:
            Me.Caption = "Eliminación de Cheques"
            btnAccion.Caption = "&Eliminar"
        Case 2:
            Me.Caption = "Anulación de Cheques"
            btnAccion.Caption = "&Anular"
    End Select
    CentraForm Me
    Show 1
    Inicio = fsComentario
End Function
Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        btnAccion.SetFocus
    End If
End Sub
Private Sub btnAccion_Click()
    If Len(Trim(txtComentario.Text)) = 0 Then
        MsgBox "Ud. debe ingresar un comentario para continuar", vbInformation, "Aviso"
        Exit Sub
    End If
    fsComentario = Trim(txtComentario.Text)
    Unload Me
End Sub
Private Sub btnSalir_Click()
    Unload Me
End Sub
