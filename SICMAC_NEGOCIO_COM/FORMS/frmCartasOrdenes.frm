VERSION 5.00
Begin VB.Form frmCartasOrdenes 
   Caption         =   "Cartas para Ordenes de Pago"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3870
   Icon            =   "frmCartasOrdenes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3870
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1380
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3705
      Begin VB.OptionButton OptAge 
         Alignment       =   1  'Right Justify
         Caption         =   "Una Agencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   510
         TabIndex        =   4
         Top             =   780
         Width           =   1995
      End
      Begin VB.OptionButton OptAge 
         Alignment       =   1  'Right Justify
         Caption         =   "Todas las Agencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   510
         TabIndex        =   3
         Top             =   330
         Width           =   2010
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   420
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "frmCartasOrdenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
    Dim clsprevio As previo.clsprevio
    Set clsprevio = New previo.clsprevio
End Sub

Private Sub cmdSalir_Click()
 Unload Me
 
End Sub
