VERSION 5.00
Begin VB.Form frmCredFormEvalDescuento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descuentos"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTotalDescuento 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptarDescuento 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitarDescuento 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregarDescuento 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin SICMACT.FlexEdit feDescuento 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   3201
      Cols0           =   4
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "N-Descripcion-Monto-Aux"
      EncabezadosAnchos=   "400-3000-1400-0"
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1-2-X"
      ListaControles  =   "0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "R-C-R-C"
      FormatosEdit    =   "3-0-3-0"
      TextArray0      =   "N"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   6
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Total :"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   2040
      Width           =   495
   End
End
Attribute VB_Name = "frmCredFormEvalDescuento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Suma As Double
Dim total As Double
Dim vnTotal As Double
Dim MtrDescuento1 As Variant
Dim MtrDescuento2 As Variant
Dim nNum As Integer

Private Sub cmdAceptarDescuento_Click()

If feDescuento.TextMatrix(1, 1) = "" Then
    MsgBox "Ud. debe Ingresar Descuentos ", vbInformation, "Aviso"
Exit Sub
End If

  Call Calculo

If nNum = 1 Then
      ReDim MtrDescuento1(2, 0)
        For i = 1 To feDescuento.Rows - 1
           ReDim Preserve MtrDescuento1(2, i)
           MtrDescuento1(0, i) = feDescuento.TextMatrix(i, 0)
            MtrDescuento1(1, i) = feDescuento.TextMatrix(i, 1)
            MtrDescuento1(2, i) = feDescuento.TextMatrix(i, 2)
           
        Next i
Else
    ReDim MtrDescuento2(2, 0)
        For i = 1 To feDescuento.Rows - 1
            ReDim Preserve MtrDescuento2(2, i)
            MtrDescuento2(0, i) = feDescuento.TextMatrix(i, 0)
            MtrDescuento2(1, i) = feDescuento.TextMatrix(i, 1)
            MtrDescuento2(2, i) = feDescuento.TextMatrix(i, 2)
           
        Next i
End If

  Unload Me
End Sub

Private Sub cmdAgregarDescuento_Click()
    If feDescuento.Rows - 1 < 25 Then
            feDescuento.lbEditarFlex = True
            feDescuento.AdicionaFila
            feDescuento.SetFocus
            SendKeys "{Enter}"
        Else
            MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdQuitarDescuento_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feDescuento.EliminaFila (feDescuento.row)
        Call Calculo
    End If
End Sub

Private Sub feDescuento_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 2 Then
        feDescuento.TextMatrix(pnRow, pnCol) = UCase(feDescuento.TextMatrix(pnRow, pnCol))
        Call Calculo
    End If
End Sub

Public Sub Calculo()
    
    txtTotalDescuento.Text = Format(SumarCampo(feDescuento, 2), "#,##0.00")
    
    Suma = txtTotalDescuento.Text
    
    txtTotalDescuento.Text = Format(Suma, "#,##0.00")
    
    vnTotal = txtTotalDescuento.Text
End Sub

Public Sub Inicio1(ByRef psTotal As Double, ByRef pMtrDescuento1 As Variant, ByVal pN As Integer)

nNum = pN

If IsArray(pMtrDescuento1) Then
        MtrDescuento1 = pMtrDescuento1
        Call CargarGridConArray1
    Else
        vnTotal = 0
End If

Me.Show 1
psTotal = vnTotal
pMtrDescuento1 = MtrDescuento1

End Sub

Public Sub Inicio2(ByRef psTotal As Double, ByRef pMtrDescuento2 As Variant, ByVal pN As Integer)
nNum = pN

If IsArray(pMtrDescuento2) Then
        MtrDescuento2 = pMtrDescuento2
        Call CargarGridConArray2
    Else
        vnTotal = 0
End If

Me.Show 1
psTotal = vnTotal
pMtrDescuento2 = MtrDescuento2
End Sub

Private Sub CargarGridConArray1()
    Dim i As Integer
    Dim nTotal As Double

    feDescuento.lbEditarFlex = True
    For i = 1 To UBound(MtrDescuento1, 2)
        feDescuento.AdicionaFila
        feDescuento.TextMatrix(i, 0) = MtrDescuento1(0, i)
        feDescuento.TextMatrix(i, 1) = MtrDescuento1(1, i)
        feDescuento.TextMatrix(i, 2) = MtrDescuento1(2, i)
        'feRemBrutaTotal.TextMatrix(i + 1, 2) = MtrRembruTotal(i + 1, 3)
        nTotal = nTotal + feDescuento.TextMatrix(i, 2)
    Next i
    txtTotalDescuento.Text = Format(nTotal, "#,##0.00")
End Sub

Private Sub CargarGridConArray2()
    Dim i As Integer
    Dim nTotal As Double

    feDescuento.lbEditarFlex = True
    For i = 1 To UBound(MtrDescuento2, 2)
        feDescuento.AdicionaFila
        feDescuento.TextMatrix(i, 0) = MtrDescuento2(0, i)
        feDescuento.TextMatrix(i, 1) = MtrDescuento2(1, i)
        feDescuento.TextMatrix(i, 2) = MtrDescuento2(2, i)
        'feRemBrutaTotal.TextMatrix(i + 1, 2) = MtrRembruTotal(i + 1, 3)
        nTotal = nTotal + feDescuento.TextMatrix(i, 2)
    Next i
    txtTotalDescuento.Text = Format(nTotal, "#,##0.00")
End Sub

Private Sub Form_Load()
CentraForm Me
End Sub
