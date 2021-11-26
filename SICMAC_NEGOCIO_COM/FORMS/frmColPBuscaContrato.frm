VERSION 5.00
Begin VB.Form frmColPBuscaContrato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colocaciones Pignoraticio - Busca Contrato"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.ListBox lstContratos 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmColPBuscaContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Formulario para mostrar los creditos de una persona
Option Explicit
Dim fsContratoSelect As String

Public Sub CargaListaContratos(ByVal psCodPers As String, ByVal psEstados As String, Optional ByVal pbSoloCreditos As Boolean = False)

Dim loPersCont As COMDColocPig.DCOMColPFunciones
Dim lrContratos As ADODB.Recordset

'On Error GoTo ControlError

    Set loPersCont = New COMDColocPig.DCOMColPFunciones
        Set lrContratos = loPersCont.dObtieneContratosPersona(psCodPers, psEstados, pbSoloCreditos)
    Set loPersCont = Nothing
    lstContratos.Clear
    If (lrContratos.BOF Or lrContratos.EOF) Then
        MsgBox " Cliente no posee Contratos ", vbInformation, " Aviso "
        fsContratoSelect = ""
        'cmdcancelar_Click
    Else
        
        Do While Not lrContratos.EOF
            lstContratos.AddItem lrContratos!cCtaCod & Space(50) & lrContratos!nPrdEstado
            lrContratos.MoveNext
        Loop
    End If
    lrContratos.Close
    Set lrContratos = Nothing
   ' lstContratos.SetFocus

End Sub

Public Sub CargaListaCtas(ByVal psCodPers As String, ByVal psEstados As String)

Dim loPersCont As COMDColocPig.DCOMColPFunciones
Dim lrContratos As ADODB.Recordset

'On Error GoTo ControlError
    
    Set loPersCont = New COMDColocPig.DCOMColPFunciones
        Set lrContratos = loPersCont.dObtieneCuentasPersona(psCodPers, psEstados, IIf(Mid(FrmMantCargoAutomatico.ActxCta.NroCuenta, 9, 1) = "1", "1", "2"))
    Set loPersCont = Nothing
     lstContratos.Clear
    If (lrContratos.BOF Or lrContratos.EOF) Then
        MsgBox " Cliente no posee Contratos ", vbInformation, " Aviso "
        fsContratoSelect = ""
        'cmdcancelar_Click
    Else
       
        Do While Not lrContratos.EOF
            lstContratos.AddItem lrContratos!cCtaCod & Space(50) & lrContratos!cPersNombre
            lrContratos.MoveNext
        Loop
    End If
    lrContratos.Close
    Set lrContratos = Nothing
   ' lstContratos.SetFocus

End Sub

Private Sub CmdAceptar_Click()
    fsContratoSelect = Mid(Me.lstContratos.Text, 1, 18)
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    fsContratoSelect = ""
End Sub

Private Sub lstContratos_DblClick()
    fsContratoSelect = Mid(Me.lstContratos.Text, 1, 18)
    Me.Hide
    'Me.cmdAceptar_Click
End Sub
