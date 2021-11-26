VERSION 5.00
Begin VB.Form frmNIIFNotasEstadoConfigDetTexto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas Estado Descripción"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7020
   Icon            =   "frmNIIFNotasEstadoConfigDetTexto.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      Height          =   4455
      Left            =   40
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   40
      Width           =   6920
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   5020
      TabIndex        =   1
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   6000
      TabIndex        =   2
      Top             =   4560
      Width           =   975
   End
End
Attribute VB_Name = "frmNIIFNotasEstadoConfigDetTexto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmNIIFNotasEstadoConfigDetTexto
'** Descripción : Configuración del Reporte Notas Estado creado segun ERS052-2013
'** Creación : EJVG, 20130425 05:40:00 PM
'********************************************************************
Option Explicit
Dim fsDescripcion As String

Public Function Inicio(ByVal psDescripcion As String) As String
    fsDescripcion = psDescripcion
    txtDescripcion.Text = fsDescripcion
    Show 1
    Inicio = fsDescripcion
End Function
Private Sub cmdAceptar_Click()
    fsDescripcion = txtDescripcion.Text
    Unload Me
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    CentraForm Me
End Sub
