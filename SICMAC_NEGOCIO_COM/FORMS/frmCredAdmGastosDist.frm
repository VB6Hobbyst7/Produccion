VERSION 5.00
Begin VB.Form frmCredAdmGastosDist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Distribucion Gastos"
   ClientHeight    =   3330
   ClientLeft      =   3765
   ClientTop       =   4545
   ClientWidth     =   6930
   Icon            =   "frmCredAdmGastosDist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   720
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   4155
      Begin VB.OptionButton OptDist 
         Caption         =   "Configurar"
         Height          =   300
         Index           =   2
         Left            =   2925
         TabIndex        =   3
         Top             =   270
         Width           =   1155
      End
      Begin VB.OptionButton OptDist 
         Caption         =   "Repartir"
         Height          =   300
         Index           =   1
         Left            =   1785
         TabIndex        =   2
         Top             =   270
         Width           =   1155
      End
      Begin VB.OptionButton OptDist 
         Caption         =   "Al Desembolso"
         Height          =   300
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   255
         Value           =   -1  'True
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCredAdmGastosDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
