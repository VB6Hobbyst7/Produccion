VERSION 5.00
Begin VB.Form frmCredSimPagosGastos 
   Caption         =   "Detalle de Gastos"
   ClientHeight    =   4365
   ClientLeft      =   3045
   ClientTop       =   3210
   ClientWidth     =   8640
   Icon            =   "frmCredSimPagosGastos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   8640
   Begin SICMACT.FlexEdit FEGastos 
      Height          =   3240
      Left            =   150
      TabIndex        =   1
      Top             =   240
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   5715
      Cols0           =   3
      FixedCols       =   0
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "-Gasto-Monto"
      EncabezadosAnchos=   "400-6000-1500"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X"
      TextStyleFixed  =   1
      ListaControles  =   "0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483628
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-R"
      FormatosEdit    =   "0-1-2"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   510
      Left            =   3285
      TabIndex        =   0
      Top             =   3660
      Width           =   1965
   End
End
Attribute VB_Name = "frmCredSimPagosGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Mostrar(ByVal pMatGastos As Variant, ByVal nNumReg As Integer)
Dim i As Integer

    For i = 0 To nNumReg - 1
        If i <> 0 Then
            FEGastos.AdicionaFila
        End If
        FEGastos.TextMatrix(i + 1, 0) = i + 1
        FEGastos.TextMatrix(i + 1, 1) = pMatGastos(i, 0)
        FEGastos.TextMatrix(i + 1, 2) = pMatGastos(i, 1)
    Next i
    Me.Show 1
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
End Sub
