VERSION 5.00
Begin VB.Form frmMensajeMostrar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado de Conexión"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblMensaje 
      Caption         =   "lblMensaje"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frmMensajeMostrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina         :   frmMensajesMostrar
'***     Descripcion    :   Formulario pa mostrar el mensaje de seguridad seleccionado aleatoriamente
'***     Creado por     :   JHCU mejoras
'***     Fecha-Creación :   22/02/2021 08:20:00 AM
'*****************************************************************************************

Option Explicit
Private fnContador As Integer
Private fnTiempoEsp As Integer
Private fbSalir As Boolean
Private nTipo As Integer
Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Public Function Inicio(ByVal psMensaje As String)
lblMensaje.Caption = Trim(psMensaje)
Me.Show 1
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If fbSalir Then
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()

fnTiempoEsp = 1
fbSalir = True
cmdAceptar.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If fbSalir Then
    Cancel = 0
Else
    Cancel = 1
End If
End Sub

