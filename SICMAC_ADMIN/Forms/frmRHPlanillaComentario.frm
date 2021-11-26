VERSION 5.00
Begin VB.Form frmRHPlanillaComentario 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmRHPlanillaComentario.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2138
      TabIndex        =   1
      Top             =   1680
      Width           =   1605
   End
   Begin VB.TextBox txtComentario 
      Height          =   1605
      Left            =   0
      MaxLength       =   255
      TabIndex        =   0
      Top             =   15
      Width           =   5850
   End
End
Attribute VB_Name = "frmRHPlanillaComentario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCadena As String
Dim lsCodEmpModiComent As String
Dim lsTipoPlanilla As String
Dim lsPeriodo As String

Public Function Ini(psCodEmp As String, psNombre As String, psComentario, Optional ByVal psTipoPlanilla As String, Optional ByVal psPeriodo As String) As String
    'se aumento psTipoPlanilla y psPeriodo para modificar el numerto de cuenta. PEAC 20131119
    Dim oPla As NActualizaDatosConPlanilla '*** PEAC 20131118
    Set oPla = New NActualizaDatosConPlanilla '*** PEAC 20131118
    
    Dim oRh As DActualizaDatosRRHH
    Set oRh = New DActualizaDatosRRHH
    Dim lsCadOriginal As String
       
    lsCodEmpModiComent = psCodEmp
    lsTipoPlanilla = psTipoPlanilla
    lsPeriodo = psPeriodo
    
    Caption = psCodEmp & " " & psNombre
    
    txtComentario.Text = psComentario
    lsCadOriginal = psComentario
    
    Me.Show 1
    
    If oRh.GetPlanillaPagada(lsCodEmpModiComent, lsTipoPlanilla, lsPeriodo) Then
        MsgBox "Ya se realizó el abono a esta cuenta, no se podrá realizar cambios.", vbOKOnly + vbExclamation, "Atención"
        Ini = lsCadOriginal
    Else
        Ini = lsCadena
        oPla.ModificaNumCtaComentario lsCodEmpModiComent, lsTipoPlanilla, lsPeriodo, Trim(lsCadena)
    End If
    
    'Ini = lsCadena
End Function

Private Sub cmdAceptar_Click()
        
    lsCadena = txtComentario.Text
    Unload Me
End Sub

Private Sub txtComentario_GotFocus()
    txtComentario.SelStart = 0
    txtComentario.SelLength = 256
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub
