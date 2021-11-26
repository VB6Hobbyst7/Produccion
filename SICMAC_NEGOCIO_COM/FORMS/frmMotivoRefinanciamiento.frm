VERSION 5.00
Begin VB.Form frmMotivoRefinanciamiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Motivo de Refinanciamiento"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmMotivoRefinanciamiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   360
      Left            =   4800
      TabIndex        =   4
      Top             =   2280
      Width           =   1100
   End
   Begin VB.TextBox txtDetalleMotivoRefinanciado 
      Height          =   1455
      Left            =   750
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   720
      Width           =   5175
   End
   Begin VB.TextBox txtMotivoRefinanciado 
      Height          =   300
      Left            =   750
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Detalle:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   75
      TabIndex        =   1
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Motivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   75
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmMotivoRefinanciamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Function Inicio(ByVal cCtaCod As String)
    Dim oNCOMCredito As New COMNCredito.NCOMCredito
    Dim oRecordset As New ADODB.Recordset
    Set oRecordset = oNCOMCredito.ObtenerMotivoRefinanciamiento(cCtaCod)
    If Not (oRecordset.BOF And oRecordset.EOF) Then
        txtMotivoRefinanciado.Text = oRecordset!cConsDescripcion
        txtDetalleMotivoRefinanciado.Text = oRecordset!sMotivoDetRef
        Show 1
    End If
End Function

