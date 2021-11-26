VERSION 5.00
Begin VB.Form frmEvaluacionAutorizacionME 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Autorizaciones ME Pendientes"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Tipos de Cambio Especial ME"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin SICMACT.FlexEdit FEevaluaciones 
         Height          =   1575
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   5655
         _extentx        =   9975
         _extenty        =   2778
         cols0           =   5
         highlight       =   2
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "N°-Hora-Monto-Estado-NroMov"
         encabezadosanchos=   "450-1500-1800-1500-0"
         font            =   "frmEvaluacionAutorizacionME.frx":0000
         font            =   "frmEvaluacionAutorizacionME.frx":0028
         font            =   "frmEvaluacionAutorizacionME.frx":0050
         font            =   "frmEvaluacionAutorizacionME.frx":0078
         font            =   "frmEvaluacionAutorizacionME.frx":00A0
         fontfixed       =   "frmEvaluacionAutorizacionME.frx":00C8
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-L-L"
         formatosedit    =   "0-0-0-0-0"
         textarray0      =   "N°"
         lbformatocol    =   -1
         lbpuntero       =   -1
         lbordenacol     =   -1
         colwidth0       =   450
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Doble Click para seleccionar  "
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
End
Attribute VB_Name = "frmEvaluacionAutorizacionME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CREATED BY GIPO ERS069-2016 30-12-2016
Dim cMovNro As String

Public Function Inicia(ByVal rs As ADODB.Recordset)
  cMovNro = ""
  Dim Estado As String
  Estado = "NARANJA"
  Do While Not rs.EOF
    FEevaluaciones.AdicionaFila
    FEevaluaciones.TextMatrix(FEevaluaciones.row, 1) = Format(rs!dFechaReg, "hh:mm:ss AMPM") 'Format(rs!dFechaReg, "MM/DD/YYYY hh:mm:ss AMPM")
    FEevaluaciones.TextMatrix(FEevaluaciones.row, 2) = Format(rs!nMontoReg, "#,#0.00")
    If (rs!nEstado = 0) Then
      Estado = "PENDIENTE"
    ElseIf (rs!nEstado = 1) Then
      Estado = "PROP. ANULADA"
    ElseIf (rs!nEstado = 2) Then
      Estado = "PROP. ENVIADA"
    End If
    FEevaluaciones.TextMatrix(FEevaluaciones.row, 3) = Estado
    FEevaluaciones.TextMatrix(FEevaluaciones.row, 4) = rs!cMovNro
    rs.MoveNext
  Loop
  Me.Show 1
  Inicia = cMovNro
End Function

Private Sub FEevaluaciones_DblClick()
    cMovNro = FEevaluaciones.TextMatrix(FEevaluaciones.row, 4)
    'MsgBox "seleccionaste " & cNroMov
    Unload Me
End Sub

