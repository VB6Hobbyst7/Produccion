VERSION 5.00
Begin VB.Form frmEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busca Nombre ..."
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   Icon            =   "frmEmpresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   3825
      TabIndex        =   3
      Top             =   750
      Width           =   930
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   2865
      TabIndex        =   2
      Top             =   750
      Width           =   930
   End
   Begin SICMACT.TxtBuscar txtPersCod 
      Height          =   330
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   582
      Appearance      =   1
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TipoBusqueda    =   3
      sTitulo         =   ""
   End
   Begin VB.Label lblPersona 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   390
      Width           =   7560
   End
End
Attribute VB_Name = "frmEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsPersCod As String

Public Function Ini() As String
    Me.Show 1
    Ini = lsPersCod
End Function

Private Sub cmdAceptar_Click()
    If Me.txtPersCod.Text = "" Then
        MsgBox "Debe elegir una persona.", vbInformation, "Aviso"
    Else
        lsPersCod = Me.txtPersCod.Text
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    lsPersCod = ""
    Unload Me
End Sub

Private Sub txtPersCod_EmiteDatos()
    Me.lblPersona.Caption = Me.txtPersCod.psDescripcion
End Sub
