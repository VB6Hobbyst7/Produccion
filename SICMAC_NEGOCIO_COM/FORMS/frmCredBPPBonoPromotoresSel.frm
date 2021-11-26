VERSION 5.00
Begin VB.Form frmCredBPPBonoPromotoresSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   Icon            =   "frmCredBPPBonoPromotoresSel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "¿Que desea Exportar?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton optDetalleSel 
         Caption         =   "Detalle del Promotor seleccionado"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Detalle de todos los Promotores"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   3015
      End
      Begin VB.OptionButton optResumen 
         Caption         =   "Resumen de Promotores"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmCredBPPBonoPromotoresSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************************************
''** Nombre      : frmCredBPPBonoPromotoresSel
''** Descripción : Formulario para exportar la generación del bono promotores
''** Creación    : WIOR, 20140620 10:00:00 AM
''*****************************************************************************************************
'Option Explicit
'Private fnTipo As Integer
'Property Let Tipo(pnTipo As String)
'   fnTipo = pnTipo
'End Property
'Property Get Tipo() As String
'    Tipo = fnTipo
'End Property
'
'Private Sub CmdAceptar_Click()
'If optResumen.value Then
'    fnTipo = 1
'ElseIf optDetalleSel.value Then
'    fnTipo = 2
'Else
'    fnTipo = 3
'End If
'Unload Me
'End Sub
'
'Private Sub cmdCerrar_Click()
'fnTipo = 0
'Unload Me
'End Sub
