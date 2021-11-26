VERSION 5.00
Begin VB.Form frmArqueoPagareFaltante 
   Caption         =   "Pagares faltantes"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12810
   Icon            =   "frmArqueoPagareFaltante.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   12810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   11520
      TabIndex        =   0
      Top             =   7080
      Width           =   1095
   End
   Begin SICMACT.FlexEdit feDetalleFaltante 
      Height          =   6495
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   12303
      Cols0           =   19
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmArqueoPagareFaltante.frx":030A
      EncabezadosAnchos=   "300-2400-2000-3000-1200-1200-1200-1200-1200-1200-1200-3000-2000-1200-1200-1200-2000-2500-3500"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      TipoBusqueda    =   6
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Crédito:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblTipoCredito 
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
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmArqueoPagareFaltante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MatPagareFaltante() As Variant
Dim CodTipoCred As String
Dim DesTipoCred As String
Dim nEstado As Integer

Public Function Inicio(ByVal pMatPagareFaltante As Variant, ByVal pCodTipoCred As String, ByVal pDesTipoCred As String, pnEstado As Integer) As Variant
    MatPagareFaltante = pMatPagareFaltante
    CodTipoCred = pCodTipoCred
    DesTipoCred = pDesTipoCred
    nEstado = pnEstado
    Me.Show 1
    Inicio = MatPagareFaltante
End Function

Private Sub Form_Load()
Call CargarDatos
End Sub

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub CargarDatos()
    lblTipoCredito = DesTipoCred
    If nEstado = 1 Then
        feDetalleFaltante.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
        feDetalleFaltante.EncabezadosAlineacion = "C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C"
        feDetalleFaltante.EncabezadosAnchos = "300-2400-2000-3000-1200-1200-1200-1200-1200-1200-1200-3000-2000-1200-1200-1200-2000-2500-3500"
        feDetalleFaltante.EncabezadosNombres = "#-Crédito-Crédito antiguo-Cliente-Moneda-M.Aprobado-S.Capital-Desembolso-Tasa-Analista-Atraso-Tipo de crédito-Condición-Estado-Desembolso BN-Oficina BN-Distrito-Zona-Direccion"
        feDetalleFaltante.FormatosEdit = "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
        feDetalleFaltante.ListaControles = "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
        
        For i = 1 To UBound(MatPagareFaltante, 2)
            If MatPagareFaltante(19, i) = CodTipoCred Then
                feDetalleFaltante.AdicionaFila
                row = feDetalleFaltante.row
                feDetalleFaltante.TextMatrix(row, 1) = MatPagareFaltante(1, i)
                feDetalleFaltante.TextMatrix(row, 2) = MatPagareFaltante(2, i)
                feDetalleFaltante.TextMatrix(row, 3) = MatPagareFaltante(3, i)
                feDetalleFaltante.TextMatrix(row, 4) = MatPagareFaltante(4, i)
                feDetalleFaltante.TextMatrix(row, 5) = MatPagareFaltante(5, i)
                feDetalleFaltante.TextMatrix(row, 6) = MatPagareFaltante(6, i)
                feDetalleFaltante.TextMatrix(row, 7) = MatPagareFaltante(7, i)
                feDetalleFaltante.TextMatrix(row, 8) = MatPagareFaltante(8, i)
                feDetalleFaltante.TextMatrix(row, 9) = MatPagareFaltante(9, i)
                feDetalleFaltante.TextMatrix(row, 10) = MatPagareFaltante(10, i)
                feDetalleFaltante.TextMatrix(row, 11) = MatPagareFaltante(11, i)
                feDetalleFaltante.TextMatrix(row, 12) = MatPagareFaltante(12, i)
                feDetalleFaltante.TextMatrix(row, 13) = MatPagareFaltante(13, i)
                feDetalleFaltante.TextMatrix(row, 14) = MatPagareFaltante(14, i)
                feDetalleFaltante.TextMatrix(row, 15) = MatPagareFaltante(15, i)
                feDetalleFaltante.TextMatrix(row, 16) = MatPagareFaltante(16, i)
                feDetalleFaltante.TextMatrix(row, 17) = MatPagareFaltante(17, i)
                feDetalleFaltante.TextMatrix(row, 18) = MatPagareFaltante(18, i)
            End If
        Next
    End If
    If nEstado = 0 Then
        feDetalleFaltante.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
        feDetalleFaltante.EncabezadosAlineacion = "C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C"
        feDetalleFaltante.EncabezadosAnchos = "300-2400-2000-1200-1200-3000-3000-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200"
        feDetalleFaltante.EncabezadosNombres = "#-Crédito-Crédito antiguo-Condición-Estado-Cliente-Producto-Desembolso-Nro C-Plazo-Vigencia-Analista-Calif. Interna-Nº Entidades-Directo Soles-Directo Dolares-Indirecto Soles-Indirecto Dolares-Calif. Sist. Finan."
        feDetalleFaltante.FormatosEdit = "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
        feDetalleFaltante.ListaControles = "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
        
        For i = 1 To UBound(MatPagareFaltante, 2)
            If MatPagareFaltante(19, i) = CodTipoCred Then
                feDetalleFaltante.AdicionaFila
                row = feDetalleFaltante.row
                feDetalleFaltante.TextMatrix(row, 1) = MatPagareFaltante(1, i)
                feDetalleFaltante.TextMatrix(row, 2) = MatPagareFaltante(2, i)
                feDetalleFaltante.TextMatrix(row, 3) = MatPagareFaltante(3, i)
                feDetalleFaltante.TextMatrix(row, 4) = MatPagareFaltante(4, i)
                feDetalleFaltante.TextMatrix(row, 5) = MatPagareFaltante(5, i)
                feDetalleFaltante.TextMatrix(row, 6) = MatPagareFaltante(6, i)
                feDetalleFaltante.TextMatrix(row, 7) = MatPagareFaltante(7, i)
                feDetalleFaltante.TextMatrix(row, 8) = MatPagareFaltante(8, i)
                feDetalleFaltante.TextMatrix(row, 9) = MatPagareFaltante(9, i)
                feDetalleFaltante.TextMatrix(row, 10) = MatPagareFaltante(10, i)
                feDetalleFaltante.TextMatrix(row, 11) = MatPagareFaltante(11, i)
                feDetalleFaltante.TextMatrix(row, 12) = MatPagareFaltante(12, i)
                feDetalleFaltante.TextMatrix(row, 13) = MatPagareFaltante(13, i)
                feDetalleFaltante.TextMatrix(row, 14) = MatPagareFaltante(14, i)
                feDetalleFaltante.TextMatrix(row, 15) = MatPagareFaltante(15, i)
                feDetalleFaltante.TextMatrix(row, 16) = MatPagareFaltante(16, i)
                feDetalleFaltante.TextMatrix(row, 17) = MatPagareFaltante(17, i)
                feDetalleFaltante.TextMatrix(row, 18) = MatPagareFaltante(18, i)
            End If
        Next
    End If
End Sub
