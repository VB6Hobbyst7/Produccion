VERSION 5.00
Begin VB.Form frmMantOpeCta 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   Icon            =   "frmMantOpeCta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2370
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   7575
   End
End
Attribute VB_Name = "frmMantOpeCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub
