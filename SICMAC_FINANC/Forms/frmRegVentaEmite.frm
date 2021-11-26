VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmRegVentaEmite 
   Caption         =   "Emisión de Facturas"
   ClientHeight    =   5625
   ClientLeft      =   1350
   ClientTop       =   2835
   ClientWidth     =   11505
   Icon            =   "frmRegVentaEmite.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   11505
   Begin VB.CheckBox chkTipoDescri 
      Caption         =   "Descripción Resumida"
      Height          =   315
      Left            =   7680
      TabIndex        =   35
      Top             =   960
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
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
      Height          =   615
      Left            =   7560
      TabIndex        =   32
      Top             =   120
      Width           =   2205
      Begin VB.OptionButton Option2 
         Caption         =   "Dolar"
         Height          =   315
         Left            =   1200
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Soles"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdBorraLinea 
      Caption         =   "Borra Linea"
      Height          =   375
      Left            =   8520
      TabIndex        =   26
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   330
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   8
      Tag             =   "txtPrincipal"
      Top             =   2640
      Width           =   5595
   End
   Begin VB.TextBox txtPreUnitario 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7080
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Añadir"
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin Sicmact.FlexEdit FlexDetalle 
      Height          =   2415
      Left            =   120
      TabIndex        =   25
      Top             =   3120
      Width           =   8295
      _extentx        =   14631
      _extenty        =   4260
      cols0           =   5
      highlight       =   1
      allowuserresizing=   3
      encabezadosnombres=   "#-Cantidad-Descripción-P.Unitario-Valor Venta"
      encabezadosanchos=   "400-1000-4000-1200-1200"
      font            =   "frmRegVentaEmite.frx":030A
      fontfixed       =   "frmRegVentaEmite.frx":0336
      columnasaeditar =   "X-X-X-X-X"
      listacontroles  =   "0-0-0-0-0"
      encabezadosalineacion=   "C-R-L-R-R"
      formatosedit    =   "0-2-0-2-2"
      textarray0      =   "#"
      lbultimainstancia=   -1  'True
      lbpuntero       =   -1  'True
      colwidth0       =   405
      rowheight0      =   300
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   345
      Left            =   8640
      TabIndex        =   24
      Top             =   5160
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   10080
      TabIndex        =   14
      Top             =   5160
      Width           =   1365
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cliente"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   7365
      Begin VB.TextBox txtNumRuc 
         Height          =   345
         Left            =   2520
         TabIndex        =   6
         Top             =   960
         Width           =   4605
      End
      Begin VB.TextBox txtDire 
         Height          =   345
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Width           =   4605
      End
      Begin Sicmact.TxtBuscar txtProvCod 
         Height          =   345
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   1935
         _extentx        =   3413
         _extenty        =   609
         appearance      =   1
         font            =   "frmRegVentaEmite.frx":0364
         tipobusqueda    =   3
         stitulo         =   ""
         tipobuspers     =   1
      End
      Begin VB.CommandButton cmdExaCab 
         Caption         =   "..."
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2250
         TabIndex        =   19
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txtProvNom 
         Height          =   345
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   4605
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Num.Doc."
         Height          =   195
         Left            =   1680
         TabIndex        =   28
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Left            =   1680
         TabIndex        =   27
         Top             =   600
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Documento"
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
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   90
      Width           =   7365
      Begin VB.TextBox txtDocSerie 
         Height          =   315
         Left            =   780
         MaxLength       =   3
         TabIndex        =   0
         Top             =   270
         Width           =   555
      End
      Begin VB.TextBox txtDocNro 
         Height          =   315
         Left            =   1350
         MaxLength       =   8
         TabIndex        =   1
         Top             =   270
         Width           =   1365
      End
      Begin MSMask.MaskEdBox txtDocFecha 
         Height          =   315
         Left            =   3420
         TabIndex        =   2
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         Height          =   225
         Left            =   2850
         TabIndex        =   17
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Número"
         Height          =   225
         Left            =   90
         TabIndex        =   16
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1185
      Left            =   8520
      TabIndex        =   20
      Top             =   3960
      Width           =   2955
      Begin VB.TextBox txtPVenta 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Height          =   285
         Left            =   1500
         TabIndex        =   13
         Top             =   840
         Width           =   1365
      End
      Begin VB.TextBox txtIGV 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   1500
         TabIndex        =   12
         Top             =   510
         Width           =   1365
      End
      Begin VB.TextBox txtVVenta 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   1500
         TabIndex        =   11
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Precio Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   23
         Top             =   870
         Width           =   1245
      End
      Begin VB.Label lblSTot 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valor Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   22
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "I.G.V."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   21
         Top             =   540
         Width           =   915
      End
      Begin VB.Shape ShapeIGV 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   60
         Top             =   120
         Width           =   2835
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   60
         Top             =   480
         Width           =   2835
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   60
         Top             =   810
         Width           =   2835
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Precio Unitario"
      Height          =   195
      Left            =   7080
      TabIndex        =   31
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Descripción"
      Height          =   195
      Left            =   1440
      TabIndex        =   30
      Top             =   2400
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad"
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   2400
      Width           =   630
   End
End
Attribute VB_Name = "frmRegVentaEmite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql   As String
Dim rs     As New ADODB.Recordset
Dim lNuevo As Boolean
Dim nTasaIGV As Currency
Dim oReg As DRegVenta
Dim sNumDoc As String
Dim lsDocRef As String
Dim lsDocFecRef As Date
'Para Ref Comercial
Dim cmdPersRefComercialEjecutado As Integer
Dim FERefComPersNoMoverdeFila As Integer
Dim lnNumRefCom As Integer
Dim MatrixHojaEval() As String
Dim nPos As Integer
Dim nDat As Integer
Dim i As Integer
Dim nTotSubTotal As Currency, nTotSubIgv As Currency, nTotFinal As Currency

Public Sub Inicio(plNuevo As Boolean, pnTasaIgv As Currency)
lNuevo = plNuevo
nTasaIGV = pnTasaIgv
Me.Show 1
End Sub

Private Function datosOk() As Boolean
datosOk = False
If txtDocSerie = "" Then
   MsgBox "Serie de Documento no definido...!", vbInformation, "! Aviso !"
   txtDocSerie.SetFocus
   Exit Function
End If
If txtDocNro = "" Then
   MsgBox "Número de Documento no definido...!", vbInformation, "! Aviso !"
   txtDocNro.SetFocus
   Exit Function
End If
If ValidaFecha(txtDocFecha) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "! Aviso !"
   txtDocFecha.SetFocus
   Exit Function
End If
If txtProvCod.Text = "" Then
   MsgBox "Proveedor no identificado...!", vbInformation, "! Aviso !"
   txtProvCod.SetFocus
   Exit Function
End If
If nVal(txtPVenta) = 0 Then
   MsgBox "Monto de Operación no indicado...!", vbInformation, "! Aviso !"
   txtVVenta.SetFocus
   Exit Function
End If
datosOk = True
End Function

Private Sub chkTipoDescri_Click()

If Me.chkTipoDescri.value = 1 Then
    Me.Label5.Visible = False
    Me.Label8.Visible = False
    Me.txtCantidad.Visible = False
    Me.txtPreUnitario.Visible = False
Else
    Me.Label5.Visible = True
    Me.Label8.Visible = True
    Me.txtCantidad.Visible = True
    Me.txtPreUnitario.Visible = True
End If

End Sub

Private Sub cmdAgregar_Click()
    Dim j As Integer
    Dim i As Integer
    Dim NCaDAr As Integer
       
If FlexDetalle.Rows > 10 Then
    MsgBox "Solo se permite diez (10) líneas...", vbInformation, "Aviso"
    Exit Sub
End If
    
If Me.chkTipoDescri.value = 1 Then
    If Len(Trim(Me.txtDescripcion.Text)) = 0 Then
        MsgBox "Ingrese Datos Correctamente...", vbInformation, "Aviso"
        Exit Sub
    End If
Else
    If Val(Me.txtCantidad.Text) = 0 Or Val(Me.txtPreUnitario.Text) = 0 Or Len(Trim(Me.txtDescripcion.Text)) = 0 Then
        MsgBox "Ingrese Datos Correctamente...", vbInformation, "Aviso"
        Exit Sub
    End If
End If

'    nDat = 1
    Me.FlexDetalle.AdicionaFila
    If Me.FlexDetalle.Row = 1 Then
        ReDim MatrixHojaEval(1 To 4, 0 To 0)
     End If
     nPos = FlexDetalle.Row - 1
     MatrixHojaEval(1, nPos) = FlexDetalle.Row
     ReDim Preserve MatrixHojaEval(1 To 4, 0 To UBound(MatrixHojaEval, 2) + 1)
     
    If nPos >= 1 Then
        
        i = nPos
        For j = 0 To i - 1
                MatrixHojaEval(1, j) = MatrixHojaEval(1, j)
                MatrixHojaEval(2, j) = MatrixHojaEval(2, j)
                MatrixHojaEval(3, j) = MatrixHojaEval(3, j)
                MatrixHojaEval(4, j) = MatrixHojaEval(4, j)

         Next j
            If Me.chkTipoDescri.value = 1 Then
                MatrixHojaEval(1, i) = Format(0, "#0.00")
                MatrixHojaEval(2, i) = Me.txtDescripcion.Text
                MatrixHojaEval(3, i) = Format(0, "#0.00")
                MatrixHojaEval(4, i) = Format(0, "#0.00")
            Else
                MatrixHojaEval(1, i) = Me.txtCantidad.Text
                MatrixHojaEval(2, i) = Me.txtDescripcion.Text
                MatrixHojaEval(3, i) = Me.txtPreUnitario.Text
                MatrixHojaEval(4, i) = Format(Me.txtCantidad.Text * Me.txtPreUnitario.Text, "#0.00")
            End If
    Else
            If Me.chkTipoDescri.value = 1 Then
                MatrixHojaEval(1, i) = Format(0, "#0.00")
                MatrixHojaEval(2, i) = Me.txtDescripcion.Text
                MatrixHojaEval(3, i) = Format(0, "#0.00")
                MatrixHojaEval(4, i) = Format(0, "#0.00")
            Else
                MatrixHojaEval(1, i) = Me.txtCantidad.Text
                MatrixHojaEval(2, i) = Me.txtDescripcion.Text
                MatrixHojaEval(3, i) = Me.txtPreUnitario.Text
                MatrixHojaEval(4, i) = Format(Me.txtCantidad.Text * Me.txtPreUnitario.Text, "#0.00")
            End If
    End If

    For i = 0 To nPos
        FlexDetalle.EliminaFila (1)
    Next i
    
    For i = 0 To nPos
        FlexDetalle.AdicionaFila
        FlexDetalle.TextMatrix(FlexDetalle.Row, 1) = MatrixHojaEval(1, i)
        FlexDetalle.TextMatrix(FlexDetalle.Row, 2) = MatrixHojaEval(2, i)
        FlexDetalle.TextMatrix(FlexDetalle.Row, 3) = MatrixHojaEval(3, i)
        FlexDetalle.TextMatrix(FlexDetalle.Row, 4) = MatrixHojaEval(4, i)
    Next

    nTotSubTotal = 0
    nTotSubIgv = 0
    nTotFinal = 0

    For i = 1 To nPos + 1
            nTotSubTotal = nTotSubTotal + FlexDetalle.TextMatrix(i, 4)
    Next i

    Me.txtVVenta.Text = Format(nTotSubTotal, "#,#.00")
    Me.txtIGV.Text = Format(nTotSubIgv, "#,#.00")
    Me.txtPVenta.Text = Format(nTotSubTotal + nTotSubIgv, "#,#.00")
    
    txtCantidad.Text = "0.00"
    Me.txtPreUnitario.Text = "0.00"
    If Me.chkTipoDescri.value = 1 Then
        Me.txtDescripcion.SetFocus
    Else
        Me.txtCantidad.SetFocus
    End If
End Sub

Private Sub cmdBorraLinea_Click()
    Dim nXPos As Integer
    nXPos = FlexDetalle.Row
    If nPos >= 1 Then
        If MsgBox("Esta Seguro de Eliminar este registro.", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            FlexDetalle.EliminaFila (FlexDetalle.Row)
            If nPos >= 1 Then

                Dim j As Integer
                For j = nXPos - 1 To nPos
                    MatrixHojaEval(1, j) = MatrixHojaEval(1, j + 1)
                    MatrixHojaEval(2, j) = MatrixHojaEval(2, j + 1)
                    MatrixHojaEval(3, j) = MatrixHojaEval(3, j + 1)
                    MatrixHojaEval(4, j) = MatrixHojaEval(4, j + 1)
                    
'                    nTotSubTotal = nTotSubTotal + MatrixHojaEval(4, j)
                    
                Next j
                nPos = nPos - 1
                
                
                nTotSubTotal = 0
                nTotSubIgv = 0
                nTotFinal = 0
            
                For i = 1 To nPos + 1
                        nTotSubTotal = nTotSubTotal + FlexDetalle.TextMatrix(i, 4)
                Next i
            
                Me.txtVVenta.Text = Format(nTotSubTotal, "#,#.00")
                Me.txtIGV.Text = Format(nTotSubIgv, "#,#.00")
                Me.txtPVenta.Text = Format(nTotSubTotal + nTotSubIgv, "#,#.00")
                
                txtCantidad.Text = "0.00"
                Me.txtPreUnitario.Text = "0.00"
                
                If Me.chkTipoDescri.value = 1 Then
                    Me.txtDescripcion.SetFocus
                Else
                    Me.txtCantidad.SetFocus
                End If
                
'                If Me.txtCantidad.Visible Then
'                    txtCantidad.SetFocus
'                Else
'                    Me.txtDescripcion.SetFocus
'                End If
                                
            Else
                nPos = nPos - 1
                nDat = 0
            End If
        End If
    Else
        If FlexDetalle.Row >= 1 Then
            FlexDetalle.EliminaFila (1)
        End If
        nPos = -1
        nDat = 0
    End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

'Private Sub fg_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
'
' fg.TextMatrix(pnRow, 4) = fg.TextMatrix(pnRow, 1) * fg.TextMatrix(pnRow, 3)
'
'End Sub

Private Sub cmdImprimir_Click()

Dim oImp As NContImprimir
Dim lsImpresion As String, lnMoneda As Integer
'Dim rsDetalle As ADODB.Recordset
Dim MatDatos() As String

Set oImp = New NContImprimir

If Val(Me.txtPVenta.Text) = 0 Then
    MsgBox "No tiene monto la factura...", vbInformation, "Aviso"
    Me.txtVVenta.SetFocus
    Exit Sub
End If

    If Me.FlexDetalle.TextMatrix(0, 1) = "" Then
        MsgBox "No hay datos para imprimir", vbInformation, "Aviso"
        Exit Sub
    End If
'Set rsDetalle = CargaRS()

ReDim Preserve MatDatos(1 To 9)

MatDatos(1) = Me.txtDocSerie.Text
MatDatos(2) = Me.txtDocNro.Text
MatDatos(3) = Me.txtDocFecha.Text
MatDatos(4) = Me.txtProvNom.Text
MatDatos(5) = Me.txtDire.Text
MatDatos(6) = Me.txtNumRuc.Text
MatDatos(7) = Me.txtVVenta.Text
MatDatos(8) = Me.txtIGV.Text
MatDatos(9) = Me.txtPVenta.Text

If Me.Option1.value = True Then
    lnMoneda = 1
Else
    lnMoneda = 2
End If

lsImpresion = ImprimirFactura(MatrixHojaEval, MatDatos, FlexDetalle.Rows, lnMoneda, Me.chkTipoDescri.value)
EnviaPrevio lsImpresion, "FACTURA", gnLinPage
Set oImp = Nothing

End Sub

Private Sub Form_Load()
Dim oPer As New UPersona
CentraForm Me
Set oReg = New DRegVenta
Set rs = oReg.CargaRegOperacion()

Dim oOpe As New DOperacion

CentraForm Me
If Not lNuevo Then
   Set rs = oReg.CargaRegistro(gnDocTpo, gsDocNro, gdFecha, gdFecha)
   If Not rs.EOF Then
      txtDocSerie = Mid(rs!cDocNro, 1, 3)
      txtDocNro = Mid(rs!cDocNro, 4, 12)
      txtDocFecha = Format(rs!dDocFecha, "dd/mm/yyyy")
      txtVVenta = Format(rs!nVVenta, gsFormatoNumeroView)
      txtIGV = Format(rs!nIGV, gsFormatoNumeroView)
      txtPVenta = Format(rs!nPVenta, gsFormatoNumeroView)
      oPer.ObtieneClientexCodigo rs!cPersCod
      txtProvCod.Tag = oPer.sPersCod
      txtProvNom = oPer.sPersNombre
      txtProvCod = oPer.sPersIdnroRUC
      If txtProvCod = "" Then
         txtProvCod = oPer.sPersIdnroDNI
      End If
   End If
End If
rs.Close: Set rs = Nothing
End Sub

Private Sub txtCantidad_GotFocus()
fEnfoque txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCantidad, KeyAscii, , , True)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtCantidad_LostFocus()
    If Trim(txtCantidad) = "" Then
        txtCantidad.Text = "0.00"
    Else
        txtCantidad.Text = Format(txtCantidad.Text, "#0.00")
    End If
End Sub

Private Sub txtDescripcion_GotFocus()
    fEnfoque txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtDire_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If

End Sub

Private Sub txtDocFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtDocFecha) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "! Aviso !"
      Exit Sub
   End If
   txtProvCod.SetFocus
End If
End Sub

Private Sub txtDocNro_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtDocNro = Format(txtDocNro, "00000000")
   txtDocFecha.SetFocus
End If
End Sub

Private Sub txtDocSerie_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtDocSerie = Format(txtDocSerie, "000")
   txtDocNro.SetFocus
End If
End Sub

Private Sub txtIgv_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtIGV, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   txtIGV = Format(txtIGV, gsFormatoNumeroView)
   txtPVenta = Format(nVal(txtVVenta) + nVal(txtIGV), gsFormatoNumeroView)
   txtPVenta.SetFocus
End If
End Sub

Private Sub txtNumRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If

End Sub

Private Sub txtPreUnitario_GotFocus()
fEnfoque txtPreUnitario
End Sub

Private Sub txtPreUnitario_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreUnitario, KeyAscii, , , True)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub txtPreUnitario_LostFocus()
    If Trim(txtPreUnitario) = "" Then
        txtPreUnitario.Text = "0.00"
    Else
        txtPreUnitario.Text = Format(txtPreUnitario.Text, "#0.00")
    End If
End Sub

Private Sub txtProvCod_EmiteDatos()
txtProvCod.Tag = txtProvCod.Text
txtProvNom = txtProvCod.psDescripcion
txtProvCod.Text = txtProvCod.sPersNroDoc
Me.txtDire.Text = txtProvCod.sPersDireccion
Me.txtNumRuc.Text = txtProvCod.sPersNroDoc
'Me.TxtDescrip.SetFocus
End Sub

Private Sub txtProvNom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If

End Sub

Private Sub txtPVenta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPVenta, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   txtVVenta = Format(Round(nVal(txtPVenta) / (1 + nTasaIGV), 2), gsFormatoNumeroView)
   txtIGV = Format(nVal(txtPVenta) - nVal(txtVVenta), gsFormatoNumeroView)
   txtPVenta = Format(txtPVenta, gsFormatoNumeroView)
   Me.cmdImprimir.SetFocus
End If
End Sub

Private Sub txtVVenta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtVVenta, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   txtIGV = Format(Round(nVal(txtVVenta) * nTasaIGV, 2), gsFormatoNumeroView)
   txtPVenta = Format(nVal(txtVVenta) + nVal(txtIGV), gsFormatoNumeroView)
   txtVVenta = Format(txtVVenta, gsFormatoNumeroView)
   txtIGV.SetFocus
End If
End Sub

Function CargaRS() As ADODB.Recordset
Dim rsHojEval As ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim contador As Integer
j = 0
If Len(FlexDetalle.TextMatrix(FlexDetalle.Row, 1)) = 0 Then
    MsgBox "No existe datos para Imprimir.", vbInformation, "Aviso"
    Exit Function
End If

Set rsHojEval = New ADODB.Recordset

With rsHojEval
    'Crear RecordSet
    
    .Fields.Append "nCantidad", adCurrency
    .Fields.Append "cDescripcion", adVarChar, 100
    .Fields.Append "nPrecioUnitario", adCurrency
    .Fields.Append "nValorVenta", adCurrency
    .Open
    
    'Llenar Recordset
        
        For i = 0 To nPos
                .AddNew
                .Fields("nCantidad") = MatrixHojaEval(1, i)
                .Fields("cDescripcion") = MatrixHojaEval(2, i)
                .Fields("nPrecioUnitario") = MatrixHojaEval(3, i)
                .Fields("nValorVenta") = MatrixHojaEval(4, i)
        Next i
End With
Set CargaRS = rsHojEval
End Function

