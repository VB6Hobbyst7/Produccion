VERSION 5.00
Begin VB.Form frmCredVerEntidades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Entidades"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7920
   Icon            =   "frmCredVerEntidades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Height          =   255
      Left            =   6720
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin SICMACT.FlexEdit FEVerEntidades 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _extentx        =   13573
      _extenty        =   5530
      cols0           =   9
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   ".-codigo-Entidad-Tipo-Anulación-cPersCod-cCtaCod-Documento-Cod_Edu"
      encabezadosanchos=   "0-0-4000-1800-1200-0-0-0-0"
      font            =   "frmCredVerEntidades.frx":030A
      font            =   "frmCredVerEntidades.frx":0336
      font            =   "frmCredVerEntidades.frx":0362
      font            =   "frmCredVerEntidades.frx":038E
      font            =   "frmCredVerEntidades.frx":03BA
      fontfixed       =   "frmCredVerEntidades.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1  'True
      columnasaeditar =   "X-X-X-X-4-X-X-X-X"
      listacontroles  =   "0-0-0-0-4-0-0-0-0"
      encabezadosalineacion=   "C-C-L-L-C-C-C-C-C"
      formatosedit    =   "0-0-0-0-0-0-0-0-0"
      textarray0      =   "."
      lbeditarflex    =   -1  'True
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
   Begin VB.Label Label1 
      Caption         =   "El check de una entidad significa que se presentó anulación de la cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   6975
   End
End
Attribute VB_Name = "frmCredVerEntidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCtaCod As String
Dim lsPersCod As String
Dim lsDocumento As String
Dim oDVent As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Dim rsCobertura As ADODB.Recordset
Dim lnFila As Integer
Dim lbAprobacion As Integer
Public Function VerEntidades(ByVal psCtaCod As String, ByVal psPersCod As String, ByVal psDocumento As String, Optional ByVal pbAprobacion As Integer = 0) As ADODB.Recordset
    lsCtaCod = psCtaCod
    lsPersCod = psPersCod
    lsDocumento = psDocumento
    lbAprobacion = pbAprobacion
    Call ObtenerEntidades
    Call ObtenerRs
    Me.Show 1
'    Call ObtenerRs
    Set VerEntidades = rsCobertura
    
End Function
Private Sub ObtenerRs()
    Dim i As Integer
    Set rsCobertura = New ADODB.Recordset
    
    'Crear RecordSet
     rsCobertura.Fields.Append "Codigo", adVarChar, 50
     rsCobertura.Fields.Append "Nombre", adVarChar, 50
     rsCobertura.Fields.Append "tipo", adVarChar, 50
     rsCobertura.Fields.Append "bAnulacion", adInteger
     rsCobertura.Fields.Append "cPersCod", adVarChar, 13
     rsCobertura.Fields.Append "cCtaCod", adVarChar, 18
     rsCobertura.Fields.Append "cDocumento", adVarChar, 30
     rsCobertura.Fields.Append "Cod_Edu", adVarChar, 10
     rsCobertura.Open
    If lnFila >= 1 Then
        For i = 1 To lnFila
            rsCobertura.AddNew
            rsCobertura.Fields("Codigo") = FEVerEntidades.TextMatrix(i, 1)
            rsCobertura.Fields("Nombre") = FEVerEntidades.TextMatrix(i, 2)
            rsCobertura.Fields("tipo") = FEVerEntidades.TextMatrix(i, 3)
            rsCobertura.Fields("bAnulacion") = IIf(FEVerEntidades.TextMatrix(i, 4) = "", 0, 1)
            rsCobertura.Fields("cPersCod") = FEVerEntidades.TextMatrix(i, 5)
            rsCobertura.Fields("cCtaCod") = FEVerEntidades.TextMatrix(i, 6)
            rsCobertura.Fields("cDocumento") = FEVerEntidades.TextMatrix(i, 7)
            rsCobertura.Fields("Cod_Edu") = FEVerEntidades.TextMatrix(i, 8)
        Next i
        rsCobertura.MoveFirst
    End If
End Sub
Private Sub ObtenerEntidades()
    lnFila = 0
    If lbAprobacion = 1 Then
        FEVerEntidades.ColumnasAEditar = "X-X-X-X-X-X-X-X-X"
    End If
    Set oDVent = New COMDCredito.DCOMCredito
    Set rs = oDVent.RecuperaVerEntidades(lsCtaCod, lsPersCod, lsDocumento)
    Set oDVent = Nothing
    Call LimpiaFlex(FEVerEntidades)
    If Not rs.EOF Then
        Do While Not rs.EOF
            FEVerEntidades.AdicionaFila
            lnFila = FEVerEntidades.row
            FEVerEntidades.TextMatrix(lnFila, 0) = "0"
            FEVerEntidades.TextMatrix(lnFila, 1) = rs!codigo
            FEVerEntidades.TextMatrix(lnFila, 2) = rs!Nombre
            FEVerEntidades.TextMatrix(lnFila, 3) = rs!Tipo
            FEVerEntidades.TextMatrix(lnFila, 4) = rs!bAnulacion
            FEVerEntidades.TextMatrix(lnFila, 5) = rs!cPersCod
            FEVerEntidades.TextMatrix(lnFila, 6) = rs!cCtaCod
            FEVerEntidades.TextMatrix(lnFila, 7) = rs!cDocumento
            FEVerEntidades.TextMatrix(lnFila, 8) = rs!Cod_Edu
            rs.MoveNext
        Loop
        FEVerEntidades.TopRow = 1
        FEVerEntidades.row = 1
    Else
        FEVerEntidades.AdicionaFila
    End If
End Sub

Private Sub FEVerEntidades_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If FEVerEntidades.Col = 4 Then
        Call ObtenerRs
    End If
End Sub

Private Sub FEVerEntidades_RowColChange()
    If FEVerEntidades.Col = 4 Then
        Call ObtenerRs
    End If
End Sub
