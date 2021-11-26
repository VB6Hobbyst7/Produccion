VERSION 5.00
Begin VB.Form frmSelecAgencias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agencias"
   ClientHeight    =   3810
   ClientLeft      =   3975
   ClientTop       =   2430
   ClientWidth     =   3570
   Icon            =   "frmSelecAgencias.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton OptSelect 
      Caption         =   "&Todos"
      Height          =   255
      Index           =   1
      Left            =   1410
      TabIndex        =   2
      Top             =   135
      Width           =   1050
   End
   Begin VB.OptionButton OptSelect 
      Caption         =   "&Ninguno"
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   1050
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   3405
   End
End
Attribute VB_Name = "frmSelecAgencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
