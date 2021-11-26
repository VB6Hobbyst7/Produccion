VERSION 5.00
Begin VB.Form FrmCredLineaCreditoLista 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Líneas de Créditos Ingresados"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14310
   Icon            =   "FrmCredLineaCreditoLista.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   14310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox txtBuscar 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nueva línea"
      Height          =   375
      Left            =   11805
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   13080
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
   End
   Begin Sicmact.FlexEdit FELista 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   9551
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Código-Demoninación-Fondo-Tasa-Estado"
      EncabezadosAnchos=   "0-2500-4000-4000-2000-1200"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-3"
      ListaControles  =   "0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-C-C"
      FormatosEdit    =   "0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label lblBuscar 
      Caption         =   "Buscar:"
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
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmCredLineaCreditoLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fbFocoGrilla As Boolean
Dim rsBuscarLineas As ADODB.Recordset 'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdNuevo_Click()
    If FrmCredLineaCredito.registro Then
         Call ListarLineaCredito
    End If
End Sub

Private Sub FELista_DblClick()
If FELista.TextMatrix(1, 0) <> "" Then
    If FrmCredLineaCredito.mantenimiento(2, Trim(FELista.TextMatrix(FELista.Row, 1))) Then
        Call ListarLineaCredito
        txtBuscar.Text = "" 'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
    End If
End If
End Sub

Private Sub FELista_GotFocus()
    fbFocoGrilla = True
End Sub

Private Sub FELista_LostFocus()
    fbFocoGrilla = False
End Sub

Private Sub Form_Load()
    txtBuscar.Text = "" 'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
    Call ListarLineaCredito
End Sub

Private Sub ListarLineaCredito()
Dim objLinea As DLineaCreditoV2
Set objLinea = New DLineaCreditoV2

Dim objRs As ADODB.Recordset
Set objRs = New ADODB.Recordset
Set objRs = objLinea.ObtenerListaLineaCredito
 Dim nNumero As Integer
    nNumero = 0
    FormateaFlex FELista
    If Not (objRs.BOF Or objRs.EOF) Then
    Set rsBuscarLineas = objRs.Clone 'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
    
        Do While Not objRs.EOF
            FELista.AdicionaFila
            nNumero = nNumero + 1
            FELista.TextMatrix(nNumero, 0) = 1
            FELista.TextMatrix(nNumero, 1) = objRs!cLineaCreditoCod
            FELista.TextMatrix(nNumero, 2) = objRs!cLineaCreditoDes
            FELista.TextMatrix(nNumero, 3) = objRs!cPersNombre
            FELista.TextMatrix(nNumero, 4) = objRs!cTasa
            FELista.TextMatrix(nNumero, 5) = objRs!cEstado
            objRs.MoveNext
        Loop
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If fbFocoGrilla Then
        If KeyCode = 86 And Shift = 2 Then
            KeyCode = 10
        End If
    End If
End Sub

'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
Private Sub txtBuscar_Change()
Dim i As Integer

If Trim(txtBuscar.Text) <> "" Then
    rsBuscarLineas.Filter = " cLineaCreditoDes LIKE '*" + Trim(txtBuscar.Text) + "*'"
    
    FELista.Clear
    FELista.FormaCabecera
    Call LimpiaFlex(FELista)
    For i = 1 To rsBuscarLineas.RecordCount
        FELista.AdicionaFila
        FELista.TextMatrix(i, 0) = 1
        FELista.TextMatrix(i, 1) = rsBuscarLineas!cLineaCreditoCod
        FELista.TextMatrix(i, 2) = rsBuscarLineas!cLineaCreditoDes
        FELista.TextMatrix(i, 3) = rsBuscarLineas!cPersNombre
        FELista.TextMatrix(i, 4) = rsBuscarLineas!cTasa
        FELista.TextMatrix(i, 5) = rsBuscarLineas!cEstado
        rsBuscarLineas.MoveNext
    Next i
Else
    Call ListarLineaCredito
End If

End Sub

Private Sub cmdActualizar_Click()
    txtBuscar.Text = ""
    Call ListarLineaCredito
End Sub
'JOEP20211111 ACTA Nº 132 - Mejora en Registrar Linea de Credito
