VERSION 5.00
Begin VB.Form frmProveedorMuestraRetencion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2685
   Icon            =   "frmProveedorMuestraRetencion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMontoRetencion 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   1280
      Width           =   1020
   End
   Begin VB.TextBox txtMontoComisionAFP 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   840
      Width           =   1020
   End
   Begin VB.TextBox txtMontoSeguroAFP 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   480
      Width           =   1020
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   320
      Left            =   75
      TabIndex        =   4
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtMontoAporte 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label Label4 
      Caption         =   "Retención :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2520
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "Comisión (AFP) :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   870
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Seguro (AFP) :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Aporte :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "frmProveedorMuestraRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************
'** Nombre : frmProveedorMuestraRetencion
'** Descripción : Formulario para mostrar la retención que se está cobrando
'** Creación : EJVG, 20140727 12:59:00 PM
'**************************************************************************
Option Explicit

Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Public Sub Iniciar(ByVal pnMontoAporte As Currency, ByVal pnMontoSeguroAFP As Currency, ByVal pnMontoComisionAFP As Currency)
    txtMontoAporte.Text = Format(pnMontoAporte, gsFormatoNumeroView)
    txtMontoSeguroAFP.Text = Format(pnMontoSeguroAFP, gsFormatoNumeroView)
    txtMontoComisionAFP.Text = Format(pnMontoComisionAFP, gsFormatoNumeroView)
    txtMontoRetencion.Text = Format(pnMontoAporte + pnMontoSeguroAFP + pnMontoComisionAFP, gsFormatoNumeroView)
    Show 1
End Sub
