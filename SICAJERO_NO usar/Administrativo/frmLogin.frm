VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acceso al Sistema"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":030A
   ScaleHeight     =   4155
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   1538
      TabIndex        =   2
      Top             =   2160
      Width           =   3165
      Begin VB.TextBox TxtClave 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1140
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   645
         Width           =   1650
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   6
         Top             =   675
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   5
         Top             =   262
         Width           =   795
      End
      Begin VB.Label LblUsu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NSSE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1155
         TabIndex        =   4
         Top             =   255
         Width           =   1620
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   405
      Left            =   3173
      TabIndex        =   1
      Top             =   3345
      Width           =   1260
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   405
      Left            =   1853
      TabIndex        =   0
      Top             =   3345
      Width           =   1275
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    If Len(Trim(Me.TxtClave.Text)) = 0 Then
        MsgBox ("Clave No puede ser en Blanco")
        Me.TxtClave.Text = ""
        Me.TxtClave.SetFocus
        Exit Sub
    End If
    If Not ClaveIncorrectaNT(Me.LblUsu.Caption, Me.TxtClave.Text, gsDominio) Then
        MsgBox ("Clave Incorrecta")
        Me.TxtClave.Text = ""
        Me.TxtClave.SetFocus
        Exit Sub
    Else
        gsPass = Trim(TxtClave.Text)
        Unload Me
        MDIMenu.Show
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim sCadTemp As String


  LblUsu.Caption = UCase(ObtenerUsuarioCliente)
  gsCodUser = LblUsu.Caption
  'gsDominio = "SOLYDES"
  gsDominio = "CMACMAYNAS"
  
End Sub

Private Sub TxtClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAceptar_Click
    End If
End Sub
