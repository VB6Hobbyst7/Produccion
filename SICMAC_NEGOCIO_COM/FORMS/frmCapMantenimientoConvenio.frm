VERSION 5.00
Begin VB.Form frmCapMantenimientoConvenio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Convenios de la Persona"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   Icon            =   "frmCapMantenimientoConvenio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2520
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2535
      Width           =   1000
   End
   Begin VB.Frame fraConvenios 
      Caption         =   "Convenio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2505
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.ListBox lstConvenios 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   130
         TabIndex        =   1
         Top             =   300
         Width           =   3240
      End
   End
End
Attribute VB_Name = "frmCapMantenimientoConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCapMantenimientoConvenio
'*** Descripción : Formulario para devolver el formulario.
'*** Creación : ELRO el 20130701 04:14:23 PM, según RFC1306270002
'********************************************************************
Option Explicit

Dim fnIdSerPag As Long

Public Function iniciarFormulario() As Long
If lstConvenios.ListCount > 0 Then
    lstConvenios.ListIndex = 0
    Me.Show 1
    iniciarFormulario = fnIdSerPag
Else
    iniciarFormulario = 0
    MsgBox "La Persona no tiene convenio.", vbInformation, "Aviso"
End If
End Function

Private Sub cmdAceptar_Click()
Dim lsCadena As String
lsCadena = lstConvenios.List(lstConvenios.ListIndex)
If Trim(lsCadena) <> "" Then
    fnIdSerPag = CInt(Trim(Right(lsCadena, 10)))
End If
Unload Me
End Sub

Private Sub cmdCancelar_Click()
fnIdSerPag = 0
Unload Me
End Sub
