VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmCFcarga 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartas Fianzas"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   4695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
End
Attribute VB_Name = "FrmCFcarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CodPersona As String
Private Sub Marco()
MSH.ClearStructure
With MSH
    .TextMatrix(0, 0) = "Carta Fianza"
    .TextMatrix(0, 1) = "Relacion"
    .TextMatrix(0, 2) = "Estado"
    
    .ColWidth(0, 0) = 2000
    .ColWidth(0, 1) = 2100
    .ColWidth(0, 2) = 2100
End With

End Sub
Public Sub Inicio(ByRef cPer As String)
CodPersona = cPer
End Sub
Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Dim Persona As UPersona

Dim CF As NCartaFianzaValida
Dim Rs As ADODB.Recordset
Set CF = New NCartaFianzaValida
Set Rs = New ADODB.Recordset
Set Rs = CF.RecuperaRelacinesInternas(CodPersona)
Dim i As Integer
Marco
i = 1
With MSH
  While Not Rs.EOF
        .AddItem i, i
        .TextMatrix(i, 0) = IIf(IsNull(Rs!Carta), "", Rs!Carta)
        .TextMatrix(i, 1) = IIf(IsNull(Rs!Relacion), "", Rs!Relacion)
        .TextMatrix(i, 2) = IIf(IsNull(Rs!Estado), "", Rs!Estado)
        i = i + 1
        Rs.MoveNext
   Wend
End With
End Sub

Private Sub MSH_Click()
If MSH.Col = 0 Then
    If MSH.Row <> 0 And MSH.TextMatrix(MSH.Row, MSH.Col) <> "" Then
        FrmCFDuplicado.CodCta.CMAC = Mid(MSH.TextMatrix(MSH.Row, MSH.Col), 1, 3)
        FrmCFDuplicado.CodCta.Age = Mid(MSH.TextMatrix(MSH.Row, MSH.Col), 4, 2)
        FrmCFDuplicado.CodCta.Prod = Mid(MSH.TextMatrix(MSH.Row, MSH.Col), 6, 3)
        FrmCFDuplicado.CodCta.Cuenta = Mid(MSH.TextMatrix(MSH.Row, MSH.Col), 9, 10)
        Unload Me
    End If
End If
End Sub
