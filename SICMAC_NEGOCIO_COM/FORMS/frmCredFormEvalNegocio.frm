VERSION 5.00
Begin VB.Form frmCredFormEvalNegocio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Negocios"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6000
   Icon            =   "frmCredFormEvalNegocio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelarNegocio 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptarNegocio 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6480
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resumen"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   5775
      Begin VB.TextBox txtIngNeto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Ingreso Neto (Negocio):"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Gastos del Negocio"
      Height          =   3735
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   5775
      Begin SICMACT.FlexEdit fgGastosNegocio 
         Height          =   3375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5475
         _extentx        =   9657
         _extenty        =   5953
         cols0           =   4
         highlight       =   1
         encabezadosnombres=   "N-Concepto-Monto-Aux"
         encabezadosanchos=   "300-3000-1400-0"
         font            =   "frmCredFormEvalNegocio.frx":030A
         font            =   "frmCredFormEvalNegocio.frx":0332
         font            =   "frmCredFormEvalNegocio.frx":035A
         font            =   "frmCredFormEvalNegocio.frx":0382
         font            =   "frmCredFormEvalNegocio.frx":03AA
         fontfixed       =   "frmCredFormEvalNegocio.frx":03D2
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         tipobusqueda    =   6
         columnasaeditar =   "X-X-2-X"
         listacontroles  =   "0-0-0-0"
         encabezadosalineacion=   "C-L-R-C"
         formatosedit    =   "0-0-2-2"
         textarray0      =   "N"
         lbeditarflex    =   -1  'True
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   300
         rowheight0      =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ventas y Costos"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.TextBox txtEgrVenta 
         Height          =   285
         Left            =   4560
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtMargBruto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtIngNegocio 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Egreso por Venta :"
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Margen Bruto :"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresos del Negocio :"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCredFormEvalNegocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre      : frmCredFormEvalNegocio
'** Descripción : Formulario para evaluación de Creditos            '
'** Referencia  : ERS004-2016                                       '
'** Creación    : JOEP, 20160525 09:00:00 AM                        '
'*******************************************************************'

Option Explicit
Dim rsDatGastoNeg As ADODB.Recordset
Dim nSumaFlex As Double
Dim vnTotal As Double
Dim MtrNegocio As Variant

Dim nIngNegocio As Double
Dim nEgrVenta As Double
Dim nMagBruto As Double
Dim nIngNeto As Double
Dim cCtaCod As String
Dim i As Integer

Dim nValorOrgNegocio As Currency
Dim nValorOrgEgresos As Currency
Dim nValorOrgMrgBruto As Currency
Dim nValorOrgIngNeto As Currency


'Cargar Datos en la grilla Gastos del Negocio
Private Sub CargarFlexEdit()
    Dim lnFila As Integer
    Dim CargarFlexEdit As Boolean
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
   
   CargarFlexEdit = oNCOMFormatosEval.CargaDatosFlexEdit(7, cCtaCod, _
                                                     rsDatGastoNeg)
'Gastos Negocio
    fgGastosNegocio.Clear
    fgGastosNegocio.FormaCabecera
    'fgGastosNegocio.Rows = 2
    Call LimpiaFlex(fgGastosNegocio)
        Do While Not rsDatGastoNeg.EOF
            fgGastosNegocio.AdicionaFila
            lnFila = fgGastosNegocio.row
            fgGastosNegocio.TextMatrix(lnFila, 0) = rsDatGastoNeg!nConsValor
            fgGastosNegocio.TextMatrix(lnFila, 1) = rsDatGastoNeg!cConsDescripcion
            fgGastosNegocio.TextMatrix(lnFila, 2) = Format(rsDatGastoNeg!nMonto, "0.00")
            rsDatGastoNeg.MoveNext
            
            Select Case CInt(fgGastosNegocio.TextMatrix(fgGastosNegocio.row, 0))
            Case 8
                fgGastosNegocio.ForeColorRow vbBlack, True
            End Select
                        
        Loop
    rsDatGastoNeg.Close
    Set rsDatGastoNeg = Nothing
End Sub
'FIN Cargar Datos en la grilla Gastos del Negocio

Private Sub cmdAceptarNegocio_Click()
    
    If txtIngNegocio.Text <= 0 Or txtEgrVenta.Text <= 0 Then
        MsgBox "Ud. debe Ingresar !Ingreso de Negocio o !Egreso de Venta ", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If txtIngNeto.Text <= 0 Then
        MsgBox "El !Ingreso Neto no puede ser '0.00' ó Negativo ", vbInformation, "Aviso"
        Exit Sub
    End If
    
Call Calcular(2)

    ReDim MtrNegocio(3, 0)
        For i = 1 To fgGastosNegocio.Rows - 1
        ReDim Preserve MtrNegocio(3, i)
            MtrNegocio(0, i) = fgGastosNegocio.TextMatrix(i, 0)
            MtrNegocio(1, i) = fgGastosNegocio.TextMatrix(i, 1)
            MtrNegocio(2, i) = fgGastosNegocio.TextMatrix(i, 2)
        Next i
        
       nIngNegocio = txtIngNegocio.Text
       nEgrVenta = txtEgrVenta.Text
       nMagBruto = txtMargBruto.Text
       nIngNeto = txtIngNeto.Text
        
Unload Me
End Sub

Private Sub cmdCancelarNegocio_Click()

nIngNegocio = nValorOrgNegocio
nEgrVenta = nValorOrgEgresos
nMagBruto = nValorOrgMrgBruto
nIngNeto = nValorOrgIngNeto

Unload Me

End Sub

Private Sub fgGastosNegocio_Click()
    Select Case CInt(fgGastosNegocio.TextMatrix(fgGastosNegocio.row, 0))
      Case 8
        Me.fgGastosNegocio.ColumnasAEditar = "X-X-X"
      Case Else
        Me.fgGastosNegocio.ColumnasAEditar = "X-X-3"
    End Select
End Sub

Private Sub fgGastosNegocio_EnterCell()
    Select Case CInt(fgGastosNegocio.TextMatrix(fgGastosNegocio.row, 0))
      Case 8
        Me.fgGastosNegocio.ColumnasAEditar = "X-X-X"
      Case Else
        Me.fgGastosNegocio.ColumnasAEditar = "X-X-3"
    End Select
End Sub

Private Sub fgGastosNegocio_OnCellChange(pnRow As Long, pnCol As Long)

    If pnCol = 2 Then
        If IsNumeric(fgGastosNegocio.TextMatrix(pnRow, pnCol)) Then
            fgGastosNegocio.TextMatrix(pnRow, pnCol) = UCase(fgGastosNegocio.TextMatrix(pnRow, pnCol))
            Call Calcular(1)
            Call Calcular(2)
        Else
            fgGastosNegocio.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
        End If
    End If
    
    Select Case pnRow
    Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
        If IsNumeric(fgGastosNegocio.TextMatrix(pnRow, pnCol)) Then
           Select Case CCur(fgGastosNegocio.TextMatrix(pnRow, pnCol))
            Case Is >= 0
                Case Else
                    MsgBox "Monto mal Ingresado ", vbInformation, "Alerta"
                    fgGastosNegocio.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
                    Exit Sub
            End Select
        Else
            fgGastosNegocio.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
        End If
    End Select
End Sub

Private Sub fgGastosNegocio_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    
    Editar = Split(Me.fgGastosNegocio.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    CentraForm Me
    txtIngNegocio.Text = "0.00"
    txtEgrVenta.Text = "0.00"
    txtMargBruto.Text = "0.00"
    txtIngNeto.Text = "0.00"
End Sub

'Calculo
Public Sub Calcular(ByVal pnTipo As Integer)
    Select Case pnTipo
    Case 1:
            txtMargBruto.Text = Format(CDbl(txtIngNegocio.Text) - CDbl(txtEgrVenta.Text), "#,##0.00")
    Case 2:
            txtIngNeto.Text = Format(CDbl(txtMargBruto.Text) - SumarCampo(fgGastosNegocio, 2), "#,##0.00")
            nIngNeto = txtIngNeto.Text
    End Select
    
    Exit Sub
End Sub
'FIN Calculo

Private Sub txtIngNegocio_GotFocus()
Me.txtIngNegocio.SelLength = Len(txtIngNegocio.Text)
End Sub

Private Sub txtEgrVenta_GotFocus()
Me.txtEgrVenta.SelLength = Len(txtEgrVenta.Text)
End Sub

Private Sub txtEgrVenta_LostFocus()

If Len(Trim(txtEgrVenta.Text)) = "" Then
        txtEgrVenta.Text = "0.00"
End If
        txtEgrVenta.Text = Format(txtEgrVenta.Text, "#,##0.00")

Call Calcular(1)
Call Calcular(2)

End Sub
Private Sub txtIngNegocio_LostFocus()

If Len(Trim(txtIngNegocio.Text)) = "" Then
        txtIngNegocio.Text = "0.00"
End If
        txtIngNegocio.Text = Format(txtIngNegocio.Text, "#,##0.00")
        
Call Calcular(1)
Call Calcular(2)

End Sub

Private Sub txtIngNegocio_KeyPress(KeyAscii As Integer)

    If txtIngNegocio.Text = "" Then
     txtIngNegocio.Text = "0.00"
    Else
        KeyAscii = NumerosDecimales(txtIngNegocio, KeyAscii, 10, , True)
        If KeyAscii = 13 Then
            'SendKeys "{Tab}", True
            EnfocaControl txtEgrVenta
            nIngNegocio = txtIngNegocio.Text
         Exit Sub
        End If
    End If
End Sub

Private Sub txtEgrVenta_KeyPress(KeyAscii As Integer)
    
If txtEgrVenta.Text = "" Then
    txtEgrVenta.Text = "0.00"
Else
    KeyAscii = NumerosDecimales(txtEgrVenta, KeyAscii, 10, , True)
       If KeyAscii = 13 Then
            'SendKeys "{Tab}", True
            EnfocaControl fgGastosNegocio
            
            nEgrVenta = txtEgrVenta.Text
            nMagBruto = txtMargBruto.Text
            
            Me.fgGastosNegocio.SetFocus
            Me.fgGastosNegocio.Col = 2
            Me.fgGastosNegocio.row = 1
            'SendKeys "{F2}"
            
        Exit Sub
        End If
End If
    
End Sub

Private Sub fgGastosNegocio_RowColChange()

     If fgGastosNegocio.Col = 2 Then
        fgGastosNegocio.AvanceCeldas = Vertical
    Else
        fgGastosNegocio.AvanceCeldas = Horizontal
    End If
    
    Select Case CInt(fgGastosNegocio.TextMatrix(fgGastosNegocio.row, 0))
      Case 8
        Me.fgGastosNegocio.ColumnasAEditar = "X-X-X"
      Case Else
        Me.fgGastosNegocio.ColumnasAEditar = "X-X-3"
    End Select
    
End Sub

Public Sub Inicio(ByVal pcCtaCod As String, ByRef pnIngNegocio As Double, ByRef pnEgrVenta As Double, ByRef pnMargBruto As Double, ByRef pnIngNeto As Double, ByRef pMtrNegocio As Variant)
    
    cCtaCod = pcCtaCod
    
    If IsArray(pMtrNegocio) Then
        
        MtrNegocio = pMtrNegocio
        Call CargarGridConArray
            
        txtIngNegocio.Text = Format(pnIngNegocio, "#,##0.00")
        txtEgrVenta.Text = Format(pnEgrVenta, "#,##0.00")
        txtMargBruto.Text = Format(pnMargBruto, "#,##0.00")
        txtIngNeto.Text = Format(pnIngNeto, "#,##0.00")
        
    Else
    
         Call CargarFlexEdit
         'vnTotal = 0
         pnIngNegocio = 0
         pnEgrVenta = 0
         pnMargBruto = 0
         pnIngNeto = 0
         
    End If

    nValorOrgNegocio = pnIngNegocio
    nValorOrgEgresos = pnEgrVenta
    nValorOrgMrgBruto = pnMargBruto
    nValorOrgIngNeto = pnIngNeto

Me.Show 1

    pMtrNegocio = MtrNegocio
           
    pnIngNegocio = nIngNegocio
    pnEgrVenta = nEgrVenta
    pnMargBruto = nMagBruto
    pnIngNeto = nIngNeto
    
End Sub

Private Sub CargarGridConArray()
    Dim i As Integer

    fgGastosNegocio.lbEditarFlex = True
    For i = 1 To UBound(MtrNegocio, 2)
        fgGastosNegocio.AdicionaFila
        fgGastosNegocio.TextMatrix(i, 0) = MtrNegocio(0, i)
        fgGastosNegocio.TextMatrix(i, 1) = MtrNegocio(1, i)
        fgGastosNegocio.TextMatrix(i, 2) = MtrNegocio(2, i)
        
        Select Case CInt(fgGastosNegocio.TextMatrix(fgGastosNegocio.row, 0))
            Case 8
                fgGastosNegocio.ForeColorRow vbBlack, True
        End Select
    Next i
    
    
End Sub




