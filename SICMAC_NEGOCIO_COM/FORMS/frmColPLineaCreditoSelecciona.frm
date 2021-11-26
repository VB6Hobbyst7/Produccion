VERSION 5.00
Begin VB.Form frmColPLineaCreditoSelecciona 
   Caption         =   "Colocaciones Pignoraticio - Linea de Crédito"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3435
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   5235
      Begin VB.ComboBox cboVersion 
         Height          =   315
         Left            =   180
         TabIndex        =   16
         Text            =   "Combo7"
         Top             =   2160
         Width           =   4785
      End
      Begin VB.ComboBox cboSubTipoCredito 
         Height          =   315
         Left            =   2700
         TabIndex        =   15
         Top             =   1620
         Width           =   2355
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         Top             =   3000
         Width           =   1185
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   3000
         Width           =   1185
      End
      Begin VB.ComboBox cboTipoCredito 
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   1620
         Width           =   2445
      End
      Begin VB.ComboBox cboPlazo 
         Height          =   315
         Left            =   2700
         TabIndex        =   9
         Top             =   990
         Width           =   2355
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   960
         Width           =   2445
      End
      Begin VB.ComboBox cboSubFondo 
         Height          =   315
         Left            =   2700
         TabIndex        =   7
         Top             =   450
         Width           =   2355
      End
      Begin VB.ComboBox cboFuenteFinanciamiento 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   450
         Width           =   2445
      End
      Begin VB.Label lblLineaCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Version"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   14
         Top             =   1980
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "SubTipo Credito"
         Height          =   195
         Index           =   5
         Left            =   2700
         TabIndex        =   13
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Credito"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   6
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Plazo"
         Height          =   195
         Index           =   3
         Left            =   2700
         TabIndex        =   5
         Top             =   810
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   810
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Sub Fondo"
         Height          =   195
         Index           =   1
         Left            =   2700
         TabIndex        =   3
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Fuente Financiamiento"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   270
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmColPLineaCreditoSelecciona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
