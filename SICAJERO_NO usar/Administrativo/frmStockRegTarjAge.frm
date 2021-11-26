VERSION 5.00
Begin VB.Form frmStockRegTarjAge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registra Stock de Tarjetas de Agencia"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmStockRegTarjAge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Stock Actual"
      Height          =   810
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   5880
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   420
         Left            =   4290
         TabIndex        =   9
         Top             =   255
         Width           =   1290
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   930
         TabIndex        =   8
         Text            =   "0"
         Top             =   285
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad :"
         Height          =   255
         Left            =   135
         TabIndex        =   10
         Top             =   315
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   5850
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   435
         Left            =   4245
         TabIndex        =   6
         Top             =   225
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.ComboBox CboUsu 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   2640
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha :"
         Height          =   285
         Left            =   60
         TabIndex        =   4
         Top             =   285
         Width           =   600
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01/01/2008"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   735
         TabIndex        =   3
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario :"
         Height          =   285
         Left            =   2325
         TabIndex        =   2
         Top             =   285
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmStockRegTarjAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
