VERSION 5.00
Begin VB.Form frmCredAutorizaComen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorizar Credito - Comentario"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5145
      TabIndex        =   2
      Top             =   2835
      Width           =   1335
   End
   Begin VB.TextBox TxtCredAutoComen 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2190
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   495
      Width           =   6330
   End
   Begin VB.Label Label1 
      Caption         =   "Comentario Obligatorio para Autorizar Credito :"
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
      Left            =   150
      TabIndex        =   1
      Top             =   105
      Width           =   4005
   End
End
Attribute VB_Name = "frmCredAutorizaComen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim bNuevoCom As Boolean
'Dim psComen As String
'
'Public Sub AgregarComentario(ByRef sComen As String)
'    bNuevoCom = True
'    TxtCredAutoComen.Enabled = True
'    Me.Show 1
'    sComen = psComen
'
'End Sub
'
'Public Sub MostarComentario(ByVal sComen As String)
'    bNuevoCom = False
'    TxtCredAutoComen.Enabled = False
'    TxtCredAutoComen.Text = sComen
'    Me.Show 1
'
'End Sub
'
'Private Sub cmdSalir_Click()
'
'    If bNuevoCom Then
'        If Trim(TxtCredAutoComen.Text) = "" Then
'            MsgBox "Debe Ingresar un Comentario", vbInformation, "Aviso"
'            Exit Sub
'        End If
'        psComen = Trim(TxtCredAutoComen.Text)
'    End If
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'    Me.Icon = LoadPicture(App.path & gsRutaIcono)
'    CentraForm Me
'End Sub
