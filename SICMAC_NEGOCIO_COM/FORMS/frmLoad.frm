VERSION 5.00
Begin VB.Form frmLoad 
   ClientHeight    =   780
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "Cargando la consulta EXPERIAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'CTI6-20210503-ERS032-2019 -(Optimizar Sugerencia)
Public Sub Inicio(oform As Object, Optional isModal As Boolean = False)
    If isModal Then
        Me.Show vbModal
    Else
        Me.Show
    End If
End Sub
