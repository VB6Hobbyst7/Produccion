VERSION 5.00
Begin VB.Form frmSegMensajeMostrar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensaje de Seguridad"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   Icon            =   "frmSegMensajeMostrar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmInicio 
      Left            =   120
      Top             =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1100
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmSegMensajeMostrar.frx":030A
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblMensaje 
      Caption         =   "lblMensaje"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmSegMensajeMostrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fnContador As Integer
Private fbSalir As Boolean
Private fnTiempoEsp As Integer

Private Sub cmdAceptar_Click()
Unload Me
End Sub

Public Function Inicio(ByVal psMensaje As String)
'WIOR 20131217 *********************
If Len(Trim(psMensaje)) > 100 Then
    lblMensaje.Width = "4800"
    Me.Width = "6000"
    cmdAceptar.Left = (Me.Width / 2) - (cmdAceptar.Width / 2)
End If
'WIOR FIN ************************
lblMensaje.Caption = Trim(psMensaje) 'WIOR 20131217
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
Dim oPar As DParametro
Set oPar = New DParametro
fnTiempoEsp = oPar.RecuperaValorParametro(3211)
Set oPar = Nothing

fnContador = 0
fbSalir = False
cmdAceptar.Visible = False
tmInicio.Interval = 1000
tmInicio.Enabled = True
Me.Height = 1620
End Sub

Private Sub Form_Unload(Cancel As Integer)
If fbSalir Then
    Cancel = 0
Else
    Cancel = 1
End If
End Sub

Private Sub tmInicio_Timer()
fnContador = fnContador + 1
If fnContador = fnTiempoEsp Then
    cmdAceptar.Visible = True
    fbSalir = True
    tmInicio.Enabled = False
    Me.Height = 2145
End If
End Sub
