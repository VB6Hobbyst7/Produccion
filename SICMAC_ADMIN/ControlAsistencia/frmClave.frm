VERSION 5.00
Begin VB.Form frmClave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingrese Clave......"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2340
   Icon            =   "frmClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   2340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   630
      TabIndex        =   1
      Top             =   525
      Width           =   1200
   End
   Begin VB.TextBox txtClave 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   45
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "**************"
      Top             =   45
      Width           =   2235
   End
End
Attribute VB_Name = "frmClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lbRes As Boolean

Public Function GetClave() As Boolean
    Me.Show 1
    GetClave = lbRes
End Function

Private Sub cmdAceptar_Click()
    If Me.txtClave.Text = gsPASS Then
        lbRes = True
    Else
        lbRes = False
    End If
    Unload Me
End Sub

Private Sub txtClave_GotFocus()
    txtClave.SelStart = 0
    txtClave.SelLength = 50
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar_Click
    End If
End Sub
