VERSION 5.00
Begin VB.Form frmCredFormEvalSobregirosPrestamos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SobreGiros y Prestamos Banca"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   12600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "&Quitar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   6
      Top             =   2760
      Width           =   1170
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   3120
      Width           =   12615
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9960
         TabIndex        =   4
         Top             =   160
         Width           =   1170
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11280
         TabIndex        =   3
         Top             =   160
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12585
      Begin SICMACT.FlexEdit feSobregiros 
         Height          =   2175
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12285
         _ExtentX        =   15954
         _ExtentY        =   3836
         Cols0           =   8
         HighLight       =   1
         EncabezadosNombres=   "N-Entidad-Credito-TEA-Cuotas-Monto Cuotas-Cuotas Pend.-Aux"
         EncabezadosAnchos=   "350-3750-1800-1200-1700-1700-1700-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ColumnasAEditar =   "X-1-2-3-4-5-6-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-C-R-R-R-C"
         FormatosEdit    =   "0-0-0-2-0-2-2-2"
         TextArray0      =   "N"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
   End
   Begin SICMACT.EditMoney txtTotalPrestamo 
      Height          =   300
      Left            =   10860
      TabIndex        =   7
      Top             =   2820
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Enabled         =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10200
      TabIndex        =   8
      Top             =   2820
      Width           =   525
   End
End
Attribute VB_Name = "frmCredFormEvalSobregirosPrestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmCredFormEvalSobregirosPrestamos
'** Descripción : Formulario para evaluación de Creditos que tienen el tipo de evaluación 2
'** Referencia  : ERS004-2016
'** Creación    : LUCV, 20160525 09:00:00 AM
'**********************************************************************************************
Option Explicit
Dim fvDetalleRef() As tFormEvalDetallePasivosSobreGirosFormato5
Dim fnNroFila As Integer
Dim fnTotal As Double
Dim fvConsCod As Integer
Dim fvConsValor As Integer

Private Sub Form_Load()
CentraForm Me
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub


Public Function Inicio(ByRef pvDetalleActivoFlex() As tFormEvalDetallePasivosSobreGirosFormato5, _
                       ByRef pnTotalCeldaFlex As Double, _
                       ByRef pnConsCod As Integer, _
                       ByRef pnConsValor As Integer, ByVal psTitulo As String) As Boolean
    
    Me.Caption = psTitulo
    If UBound(pvDetalleActivoFlex) > 0 Then 'Si Matrix Contiene Datos
    fvDetalleRef = pvDetalleActivoFlex
    Call SetFlexDetalleSobreGiros
    fnTotal = pnTotalCeldaFlex
    fvConsCod = pnConsCod
    fvConsValor = pnConsValor
    Call SumarMontos
    Else
    ReDim pvDetalleActivoFlex(0)
    pnTotalCeldaFlex = 0
    fnTotal = 0
    End If
        
    Me.Show 1
    pvDetalleActivoFlex = fvDetalleRef
    pnTotalCeldaFlex = fnTotal
    pnConsCod = fvConsValor
    pnConsValor = fvConsValor
    Inicio = True
    
End Function

Private Sub SetFlexDetalleSobreGiros()
    Dim index As Integer
    feSobregiros.lbEditarFlex = True
    Call LimpiaFlex(feSobregiros)

    For index = 1 To UBound(fvDetalleRef)
            feSobregiros.AdicionaFila
            feSobregiros.TextMatrix(index, 1) = fvDetalleRef(index).cEntidad
            feSobregiros.TextMatrix(index, 2) = fvDetalleRef(index).cCredito
            feSobregiros.TextMatrix(index, 3) = Format(fvDetalleRef(index).nTEA, "#,##0.00")
            feSobregiros.TextMatrix(index, 4) = fvDetalleRef(index).nCuotas
            feSobregiros.TextMatrix(index, 5) = Format(fvDetalleRef(index).nMontoCuota, "#,##0.00")
            feSobregiros.TextMatrix(index, 6) = Format(fvDetalleRef(index).nCuotasPend, "#,##0.00")
    Next
    Call SumarMontos
End Sub
Private Sub SumarMontos()
    Dim i As Integer
    Dim lnMontoCuota As Currency
    Dim lnCuotaPend As Currency
    Dim lnTotal As Currency
    lnTotal = 0
    If feSobregiros.TextMatrix(1, 0) <> "" Then
        For i = 1 To feSobregiros.Rows - 1
            lnMontoCuota = IIf(IsNumeric(feSobregiros.TextMatrix(i, 5)), feSobregiros.TextMatrix(i, 5), 0)
            lnCuotaPend = IIf(IsNumeric(feSobregiros.TextMatrix(i, 6)), feSobregiros.TextMatrix(i, 6), 0)
            lnTotal = lnTotal + CCur(lnMontoCuota * lnCuotaPend)
        Next
    End If
    txtTotalPrestamo.Enabled = False
    txtTotalPrestamo.Text = Format(lnTotal, "#,##0.00")
    fnTotal = Format(lnTotal, "#,##0.00")
End Sub

Private Sub cmdAceptar_Click()
Dim index As Integer
Dim i As Integer
'If Not validarDetalle Then Exit Sub
If MsgBox("Desea Guardar Cuentas x Cobrar??", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'Llenado de Matriz
index = Me.feSobregiros.Rows - 1
ReDim Preserve fvDetalleRef(index)
For i = 1 To index
    fvDetalleRef(i).cEntidad = feSobregiros.TextMatrix(i, 1)
    fvDetalleRef(i).cCredito = feSobregiros.TextMatrix(i, 2)
    fvDetalleRef(i).nTEA = CDbl(feSobregiros.TextMatrix(i, 3))
    fvDetalleRef(i).nCuotas = CInt(feSobregiros.TextMatrix(i, 4))
    fvDetalleRef(i).nMontoCuota = feSobregiros.TextMatrix(i, 5)
    fvDetalleRef(i).nCuotasPend = feSobregiros.TextMatrix(i, 6)
Next i
fnTotal = CDbl(txtTotalPrestamo.Text)
Unload Me
End Sub

Private Sub cmdAgregar_Click()
    If feSobregiros.Rows - 1 < 25 Then
        feSobregiros.lbEditarFlex = True
        feSobregiros.AdicionaFila
        feSobregiros.SetFocus
        SendKeys "{Enter}"
    Else
    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdQuitar_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feSobregiros.EliminaFila (feSobregiros.row)
        txtTotalPrestamo.Text = Format(SumarCampo(feSobregiros, 3), "#,##0.00")
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub feSobregiros_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 1 Then
    feSobregiros.TextMatrix(feSobregiros.row, 2) = UCase(feSobregiros.TextMatrix(feSobregiros.row, 2))
    End If
    'txtTotalPrestamo.Text = Format(SumarCampo(feSobregiros, 5), "#,##0.00")
    Call SumarMontos
End Sub

Private Sub feSobregiros_OnRowChange(pnRow As Long, pnCol As Long)
feSobregiros.TextMatrix(feSobregiros.row, 1) = UCase(feSobregiros.TextMatrix(feSobregiros.row, 1))
End Sub
