VERSION 5.00
Begin VB.Form frmParametroRevisionExp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetro de Revisión"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmParametroRevisionExp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtMontoParExp 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblParRevExp 
      Caption         =   "Parámetro Límite para Revisión S/.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmParametroRevisionExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nParRevision As Integer

Private Sub cmdGuardar_Click()
    Dim oNCOMColocEval As NCOMColocEval
    Set oNCOMColocEval = New NCOMColocEval
    Dim oNCOMContFunciones As COMNContabilidad.NCOMContFunciones
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    Call oNCOMColocEval.updateParametroRevision(Me.txtMontoParExp.Text)
    MsgBox "Parámetro Actualizado Correctamente", vbInformation, "Aviso"
End Sub

Private Sub Form_Load()
    Dim oNCOMColocEval As NCOMColocEval
    Set oNCOMColocEval = New NCOMColocEval
    Dim rsParRevision As ADODB.Recordset
    Set rsParRevision = New ADODB.Recordset
    Dim nParRevision As Integer
    Set rsParRevision = oNCOMColocEval.obtenerParametroRevision()
    nParRevision = rsParRevision!nParValor
    Me.txtMontoParExp.Text = Format(nParRevision, "#0.00")
End Sub
