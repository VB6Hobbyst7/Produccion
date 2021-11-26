VERSION 5.00
Begin VB.Form frmCredSolicAutAmpComent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitud de exoneración de Ampliación"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "frmCredSolicAutAmpComent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtComentario 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   5535
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   2
      Top             =   600
      Width           =   1170
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese el motivo de la solicitud de exoneración:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label lblFecha 
      Caption         =   "lblFecha"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label lblAgencia 
      Caption         =   "lblAgencia"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label lblUsuario 
      Caption         =   "lblUsuario"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmCredSolicAutAmpComent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredSolicAutAmpComent
'** Descripción : Formulario para ingresar el sustento de las solicitudes con ampliaciones
'**               excepcionales creado segun TI-ERS030-2016
'** Creación : JUEZ, 20160510 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim lsComentario As String

Public Function ObtenerComentario() As String
    lblUsuario.Caption = "Usuario: " & gsCodUser
    lblAgencia.Caption = "Agencia: " & UCase(gsNomAge)
    lblFecha.Caption = "Fecha: " & Format(gdFecSis, gsFormatoFechaView)
    Me.Show 1
    ObtenerComentario = lsComentario
End Function

Private Sub cmdAceptar_Click()
    If Trim(txtComentario.Text) = "" Then
        MsgBox "Debe ingresar el motivo", vbInformation, "Aviso"
        txtComentario.SetFocus
        Exit Sub
    End If
    lsComentario = Trim(txtComentario.Text)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    lsComentario = ""
    txtComentario.Text = ""
    Unload Me
End Sub
