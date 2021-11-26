VERSION 5.00
Begin VB.Form frmCredSimNroCuotas 
   Caption         =   "Calculo de Numero de Cuotas"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4410
   Icon            =   "frmCredSimNroCuotas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   465
      Left            =   1455
      TabIndex        =   2
      Top             =   2100
      Width           =   1665
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   45
      TabIndex        =   1
      Top             =   1275
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   4320
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1305
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Prestamo  :"
         Height          =   270
         Left            =   180
         TabIndex        =   3
         Top             =   285
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmCredSimNroCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
