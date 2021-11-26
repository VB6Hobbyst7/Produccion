VERSION 5.00
Begin VB.Form frmActivarOffHost 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar OFF Host"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmActivarOffHost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   45
      TabIndex        =   1
      Top             =   1125
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "Salir"
         Height          =   315
         Left            =   3120
         TabIndex        =   3
         Top             =   210
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   990
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   4590
      Begin VB.CheckBox Check1 
         Caption         =   "ACTIVAR O DESACTIVAR HOST"
         Height          =   300
         Left            =   315
         TabIndex        =   2
         Top             =   345
         Width           =   3060
      End
   End
End
Attribute VB_Name = "frmActivarOffHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
