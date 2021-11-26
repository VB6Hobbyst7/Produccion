VERSION 5.00
Begin VB.Form frmListaVinculadosCargo 
   Caption         =   "Lista Vinculados Cargo"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin VB.ListBox lstCargos 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "frmListaVinculadosCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public psValorSel As String

Private Sub btnAceptar_Click()
    psValorSel = lstCargos.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oGrup As COMDpersona.DCOMGrupoE
    Dim oDR As ADODB.Recordset
    Dim i As Integer
    Set oGrup = New COMDpersona.DCOMGrupoE
    Set oDR = New ADODB.Recordset
    
    Set oDR = oGrup.ObtenerRelacionGestion
    Set oGrup = Nothing
    If Not (oDR.BOF And oDR.EOF) Then
        For i = 1 To oDR.RecordCount
            lstCargos.AddItem (oDR!cRelacionGestion & " - " & oDR!cDescRelacionGestion)
            oDR.MoveNext
        Next
    End If
End Sub
