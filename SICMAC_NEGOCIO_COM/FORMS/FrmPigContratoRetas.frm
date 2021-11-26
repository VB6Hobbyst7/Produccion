VERSION 5.00
Begin VB.Form FrmPigContratoRetas 
   Caption         =   "Contrato - Retasacion Manual"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin SICMACT.FlexEdit FlexEdit1 
      Height          =   3345
      Left            =   45
      TabIndex        =   2
      Top             =   630
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   5900
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   -1
      RowHeight0      =   240
      ForeColorFixed  =   -2147483630
   End
   Begin VB.TextBox txtContrato 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   780
      TabIndex        =   1
      Top             =   105
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Contrato"
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   165
      Width           =   630
   End
End
Attribute VB_Name = "FrmPigContratoRetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
