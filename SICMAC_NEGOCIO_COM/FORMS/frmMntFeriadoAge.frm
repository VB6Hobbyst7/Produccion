VERSION 5.00
Begin VB.Form frmMntFeriadoAge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agencias Relacionadas al Feriado"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   ControlBox      =   0   'False
   Icon            =   "frmMntFeriadoAge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   2580
      TabIndex        =   1
      Top             =   4860
      Width           =   1425
   End
   Begin SICMACT.FlexEdit FEAge 
      Height          =   4650
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   8202
      Cols0           =   4
      FixedCols       =   0
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "-Nro-Agencia-Valor"
      EncabezadosAnchos=   "600-600-3600-600"
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
      ListaControles  =   "0-0-0-4"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-C"
      FormatosEdit    =   "0-0-0-0"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   600
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmMntFeriadoAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MatAgenciasLocal As Variant

Public Function CargaFlex(ByRef MatAgencias As Variant, ByVal pHabilita As Boolean) As Variant
Dim i As Integer

    For i = 0 To UBound(MatAgencias) - 1
        FEAge.AdicionaFila , , True
        FEAge.TextMatrix(i + 1, 1) = MatAgencias(i, 0) 'Numero de Agencia
        FEAge.TextMatrix(i + 1, 2) = MatAgencias(i, 1) 'Descripcion de Agencia
        FEAge.TextMatrix(i + 1, 3) = MatAgencias(i, 2) 'check
        If pHabilita Then
            FEAge.TextMatrix(i + 1, 3) = "1"
        End If
    Next i
    
    MatAgenciasLocal = MatAgencias
    
    Me.FEAge.Enabled = pHabilita
    
    Me.Show 1
    MatAgencias = MatAgenciasLocal
    CargaFlex = MatAgencias
End Function

Private Sub CmdAceptar_Click()
    Dim i As Integer

    For i = 0 To UBound(MatAgenciasLocal) - 1
        MatAgenciasLocal(i, 2) = FEAge.TextMatrix(i + 1, 3) 'check
    Next i

    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
End Sub
