VERSION 5.00
Begin VB.Form frmIntangibleProvision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intangibles por Activar"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14760
   Icon            =   "frmIntangibleProvision.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   14760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   12240
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   13440
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin Sicmact.FlexEdit feDocProvision 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   9340
      Cols0           =   8
      HighLight       =   2
      AllowUserResizing=   1
      EncabezadosNombres=   "#-Proveedor-Comprobante-Fecha-Moneda-Monto-Descripcion-nMovNro"
      EncabezadosAnchos=   "300-3000-1800-1200-1200-1200-5500-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-C-R-L-R"
      FormatosEdit    =   "0-1-1-1-1-2-1-3"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColor       =   0
      CellForeColor   =   0
   End
End
Attribute VB_Name = "frmIntangibleProvision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fRsComprobante As New ADODB.Recordset
Dim fsDocOrigenNro As String
Dim fnComproMovNro As Long
Public Function Inicio(ByVal pRSComprobante As ADODB.Recordset, ByRef psDocOrigenNro As String, ByRef pnComprobMovNro As Long) As String
    Set fRsComprobante = pRSComprobante.Clone
    Show 1
    psDocOrigenNro = fsDocOrigenNro
    pnComprobMovNro = fnComproMovNro
End Function
Private Sub cmdAceptar_Click()
    If feDocProvision.TextMatrix(1, 0) = "" Then
        MsgBox "No existen comprobantes pendientes", vbInformation, "Aviso"
        Exit Sub
    End If
    fnComproMovNro = CLng(feDocProvision.TextMatrix(feDocProvision.row, 7))
    fsDocOrigenNro = feDocProvision.TextMatrix(feDocProvision.row, 2)
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    feDocProvision.SetFocus
    SendKeys "{Right}"
End Sub
Private Sub Form_Load()
    Dim row As Long
    LimpiaFlex feDocProvision
    Do While Not fRsComprobante.EOF
        feDocProvision.AdicionaFila
        row = feDocProvision.row
        feDocProvision.TextMatrix(row, 1) = fRsComprobante!Proveedor
        feDocProvision.TextMatrix(row, 2) = fRsComprobante!Comprobante
        feDocProvision.TextMatrix(row, 3) = fRsComprobante!Fecha
        feDocProvision.TextMatrix(row, 4) = fRsComprobante!Moneda
        feDocProvision.TextMatrix(row, 5) = Format(fRsComprobante!Importe, "#,#0.00")
        feDocProvision.TextMatrix(row, 6) = fRsComprobante!Glosa
        feDocProvision.TextMatrix(row, 7) = fRsComprobante!nMovNro
        fRsComprobante.MoveNext
    Loop
    If fRsComprobante.RecordCount > 0 Then
        cmdAceptar.Default = True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set fRsComprobante = Nothing
End Sub

