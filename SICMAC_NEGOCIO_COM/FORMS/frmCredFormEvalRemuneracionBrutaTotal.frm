VERSION 5.00
Begin VB.Form frmCredFormEvalRemuneracionBrutaTotal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remuneracion Bruta Total"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5280
   Icon            =   "frmCredFormEvalRemuneracionBrutaTotal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtTotalRemBrutaTotal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptarRemBrutaTotal 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitarRemBrutaTotal 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregarRemBrutaTotal 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   3480
      Width           =   975
   End
   Begin SICMACT.FlexEdit feRemBrutaTotal 
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5040
      _extentx        =   9102
      _extenty        =   4471
      cols0           =   5
      highlight       =   1
      allowuserresizing=   1
      encabezadosnombres=   "AuxN-N-Descripcion-Monto-Aux2"
      encabezadosanchos=   "0-400-3000-1400-0"
      font            =   "frmCredFormEvalRemuneracionBrutaTotal.frx":030A
      fontfixed       =   "frmCredFormEvalRemuneracionBrutaTotal.frx":0336
      columnasaeditar =   "X-X-X-3-X"
      listacontroles  =   "0-0-0-0-0"
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      encabezadosalineacion=   "C-C-L-R-C"
      formatosedit    =   "0-0-0-2-0"
      textarray0      =   "AuxN"
      lbeditarflex    =   -1  'True
      lbultimainstancia=   -1  'True
      tipobusqueda    =   6
      lbbuscaduplicadotext=   -1  'True
      rowheight0      =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Total :"
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
      Left            =   3480
      TabIndex        =   4
      Top             =   2205
      Width           =   735
   End
End
Attribute VB_Name = "frmCredFormEvalRemuneracionBrutaTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre      : frmCredFormEvalRemuneracionBrutaTotal                 '
'** Descripción : Formulario para evaluación de Creditos            '
'** Referencia  : ERS004-2016                                       '
'** Creación    : JOEP, 20160525 09:00:00 AM                        '
'*******************************************************************'

Option Explicit

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Dim rsDatRemBrutaTotalMes1 As ADODB.Recordset
Dim rsDatRemBrutaTotalMes2 As ADODB.Recordset

Dim pcCtaCod As String
Dim rsOp1 As ADODB.Recordset
Dim rsOp2 As ADODB.Recordset
Dim rsOp3 As ADODB.Recordset
Dim rsOp4 As ADODB.Recordset
Dim rsOp5 As ADODB.Recordset
Dim rsOp6 As ADODB.Recordset
Dim rsOp7 As ADODB.Recordset
Dim rsOp8 As ADODB.Recordset
Dim rsOp9 As ADODB.Recordset
Dim rsOp10 As ADODB.Recordset
Dim rsOp11 As ADODB.Recordset
Dim rsOp12 As ADODB.Recordset
Dim rsOp13 As ADODB.Recordset

Dim Suma1 As Double
Dim Suma2 As Double
Dim Total As Double
Dim vnTotal1 As Double
Dim vnTotal2 As Double
Dim MtrRembruTotal1 As Variant
Dim MtrRembruTotal2 As Variant
Dim nNum As Integer

Dim nFilaPrimero1 As Double
Dim nFilaPrimero2 As Double
Dim nTotalCompraDeu1 As Currency
Dim nTotalCompraDeu2 As Currency
Dim i As Integer

Dim nOrigPrimera1 As Currency
Dim nOrigTotal1 As Currency
Dim nOrigCompraDeu1 As Currency

Dim nOrigPrimera2 As Currency
Dim nOrigTotal2 As Currency
Dim nOrigCompraDeu2 As Currency

Dim fsCtaCod As String

Private Sub cmdAceptarRemBrutaTotal_Click()

If feRemBrutaTotal.TextMatrix(1, 3) <= 0 Then
    MsgBox "Ud. debe Ingresar Monto de Remuneracion Bruta Total", vbInformation, "Aviso"
    Exit Sub
End If

If nNum = 1 Then
    nTotalCompraDeu1 = 0
      ReDim MtrRembruTotal1(3, 0)
        For i = 1 To feRemBrutaTotal.Rows - 1
            ReDim Preserve MtrRembruTotal1(3, i)
            MtrRembruTotal1(1, i) = feRemBrutaTotal.TextMatrix(i, 1)
            MtrRembruTotal1(2, i) = feRemBrutaTotal.TextMatrix(i, 2)
            MtrRembruTotal1(3, i) = feRemBrutaTotal.TextMatrix(i, 3)
            
            If feRemBrutaTotal.TextMatrix(i, 1) = 6 Or feRemBrutaTotal.TextMatrix(i, 1) = 5 Then
                nTotalCompraDeu1 = nTotalCompraDeu1 + feRemBrutaTotal.TextMatrix(i, 3)
            End If
            
        Next i
       Call Calculo(2)
       Call Calculo(3)
Else
    nTotalCompraDeu2 = 0
    ReDim MtrRembruTotal2(3, 0)
        For i = 1 To feRemBrutaTotal.Rows - 1
            ReDim Preserve MtrRembruTotal2(3, i)
            MtrRembruTotal2(1, i) = feRemBrutaTotal.TextMatrix(i, 1)
            MtrRembruTotal2(2, i) = feRemBrutaTotal.TextMatrix(i, 2)
            MtrRembruTotal2(3, i) = feRemBrutaTotal.TextMatrix(i, 3)
            
            If feRemBrutaTotal.TextMatrix(i, 1) = 6 Or feRemBrutaTotal.TextMatrix(i, 1) = 5 Then
                nTotalCompraDeu2 = nTotalCompraDeu2 + feRemBrutaTotal.TextMatrix(i, 3)
            End If
            
        Next i
       Call Calculo(2)
       Call Calculo(3)
End If
  Unload Me
End Sub

Private Sub cmdAgregarRemBrutaTotal_Click()
    If feRemBrutaTotal.Rows - 1 < 25 Then
            feRemBrutaTotal.lbEditarFlex = True
            feRemBrutaTotal.AdicionaFila
            feRemBrutaTotal.SetFocus
            SendKeys "{Enter}"
        Else
            MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdQuitarRemBrutaTotal_Click()
    If feRemBrutaTotal.TextMatrix(feRemBrutaTotal.row, 0) = "1" Then
        MsgBox "Solo se puede Editar?", vbInformation, "Aviso"
        Exit Sub
    ElseIf MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Call Calculo(1)
        feRemBrutaTotal.EliminaFila (feRemBrutaTotal.row)
        
    End If
End Sub

Private Sub Command1_Click()
If nNum = 1 Then
    nFilaPrimero1 = nOrigPrimera1
    vnTotal1 = nOrigTotal1
    nTotalCompraDeu1 = nOrigCompraDeu1
Else
    nFilaPrimero2 = nOrigPrimera2
    vnTotal2 = nOrigTotal2
    nTotalCompraDeu2 = nOrigCompraDeu2
End If
 Unload Me

End Sub

Private Sub feRemBrutaTotal_Click()

If nNum = 2 Then
    Select Case CInt(feRemBrutaTotal.TextMatrix(feRemBrutaTotal.row, 1))
      Case 4
        'Me.feRemBrutaTotal.ColumnasAEditar = "X-X-X-X-X"
      Case Else
        Me.feRemBrutaTotal.ColumnasAEditar = "X-X-X-3-X"
    End Select
End If

End Sub

Private Sub feRemBrutaTotal_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 3 Then
        If IsNumeric(feRemBrutaTotal.TextMatrix(pnRow, pnCol)) Then
            If feRemBrutaTotal.TextMatrix(pnRow, pnCol) >= 0 Then
                If feRemBrutaTotal.TextMatrix(feRemBrutaTotal.row, 0) = "1" Then
                Else
                    feRemBrutaTotal.TextMatrix(pnRow, pnCol) = UCase(Format(feRemBrutaTotal.TextMatrix(pnRow, pnCol), "#,##0.00"))
                    Call Calculo(2)
                    Call Calculo(3)
                End If
            Else
                feRemBrutaTotal.TextMatrix(pnRow, pnCol) = "0.00"
            End If
        Else
            feRemBrutaTotal.TextMatrix(pnRow, pnCol) = "0.00"
        End If
    End If
        
End Sub

Public Sub Calculo(ByVal nTipo As Integer)
        
If nNum = 1 Then
        Select Case nTipo
            Case 1:
                 Suma1 = Suma1 - feRemBrutaTotal.TextMatrix(feRemBrutaTotal.row, 3)
                
                 txtTotalRemBrutaTotal.Text = Format(Suma1, "#,##0.00")
                 
                 vnTotal1 = txtTotalRemBrutaTotal.Text
            Case 2:
            Suma1 = 0
                For i = 2 To feRemBrutaTotal.Rows - 1
                    Suma1 = Suma1 + feRemBrutaTotal.TextMatrix(i, 3)
                Next i
                
            txtTotalRemBrutaTotal.Text = Format(Suma1, "#,##0.00")
                 
            vnTotal1 = txtTotalRemBrutaTotal.Text
            
            Case 3:
            
            If feRemBrutaTotal.TextMatrix(1, 0) = "1" Then
                    If nNum = 1 Then
                        nFilaPrimero1 = feRemBrutaTotal.TextMatrix(1, 3)
                    Else
                        nFilaPrimero2 = feRemBrutaTotal.TextMatrix(1, 3)
                    End If
            End If
                
        End Select
Else
        Select Case nTipo
            Case 1:
            Suma2 = Suma2 - feRemBrutaTotal.TextMatrix(feRemBrutaTotal.row, 3)
                
                 txtTotalRemBrutaTotal.Text = Format(Suma2, "#,##0.00")
                 
                 vnTotal2 = txtTotalRemBrutaTotal.Text
            Case 2:
            Suma2 = 0
                For i = 2 To feRemBrutaTotal.Rows - 1
                    Suma2 = Suma2 + feRemBrutaTotal.TextMatrix(i, 3)
                Next i
            
            txtTotalRemBrutaTotal.Text = Format(Suma2, "#,##0.00")
                           
            vnTotal2 = txtTotalRemBrutaTotal.Text
            
            Case 3:
            
            If feRemBrutaTotal.TextMatrix(1, 0) = "1" Then
                    If nNum = 1 Then
                        nFilaPrimero1 = feRemBrutaTotal.TextMatrix(1, 3)
                    Else
                        nFilaPrimero2 = feRemBrutaTotal.TextMatrix(1, 3)
                    End If
            End If
        
        End Select
End If
End Sub

Public Sub Inicio1(ByRef psTotal As Double, ByRef psFilaPrimero As Double, ByRef pMtrRemuBrutaTotal1 As Variant, ByVal pN As Integer, ByRef psTotalCompraDeuda As Currency, ByVal psCtaCod As String, _
                    ByVal pnTpoInsConv As Integer, ByVal pbSecSalud As Boolean)
    fsCtaCod = psCtaCod
    nNum = pN
        
    If IsArray(pMtrRemuBrutaTotal1) Then
        MtrRembruTotal1 = pMtrRemuBrutaTotal1
        Call CargarGridConArray1
        nOrigCompraDeu1 = psTotalCompraDeuda
        
        If (pnTpoInsConv = 1 And pbSecSalud = False) Or (pnTpoInsConv = 1 And pbSecSalud = True) Then
            Label1.Visible = False
            txtTotalRemBrutaTotal.Visible = False
        ElseIf pnTpoInsConv = 2 And pbSecSalud = False Then
            Label1.Visible = True
            txtTotalRemBrutaTotal.Visible = True
        End If
        
    Else
        
        Call CargarFlexEdit(pnTpoInsConv, pbSecSalud)
        
        nFilaPrimero1 = 0
        vnTotal1 = 0
        nTotalCompraDeu1 = 0
        
        nOrigPrimera1 = nFilaPrimero1
        nOrigTotal1 = vnTotal1
        nOrigCompraDeu1 = nTotalCompraDeu1
        
            If pnTpoInsConv = 1 And pbSecSalud = True Then
                feRemBrutaTotal.EliminaFila (2)
            End If
            
            If (pnTpoInsConv = 1 And pbSecSalud = False) Or (pnTpoInsConv = 1 And pbSecSalud = True) Then
                Label1.Visible = False
                txtTotalRemBrutaTotal.Visible = False
            ElseIf pnTpoInsConv = 2 And pbSecSalud = False Then
                Label1.Visible = True
                txtTotalRemBrutaTotal.Visible = True
            End If
            
    End If

Me.Show 1
If IsArray(MtrRembruTotal1) Then
    psTotal = vnTotal1
    psFilaPrimero = nFilaPrimero1
    pMtrRemuBrutaTotal1 = MtrRembruTotal1
    psTotalCompraDeuda = nTotalCompraDeu1
End If
End Sub

Public Sub Inicio2(ByRef psTotal As Double, ByRef psFilaPrimero As Double, ByRef pMtrRemuBrutaTotal2 As Variant, ByVal pN As Integer, ByRef psTotalCompraDeuda As Currency, ByVal psCtaCod As String, _
                    ByVal pnTpoInsConv As Integer, ByVal pbSecSalud As Boolean)

    fsCtaCod = psCtaCod
    
    nNum = pN
    
    If IsArray(pMtrRemuBrutaTotal2) Then
            MtrRembruTotal2 = pMtrRemuBrutaTotal2
            Call CargarGridConArray2
            nOrigCompraDeu2 = psTotalCompraDeuda
            
            If (pnTpoInsConv = 1 And pbSecSalud = False) Or (pnTpoInsConv = 1 And pbSecSalud = True) Then
                Label1.Visible = False
                txtTotalRemBrutaTotal.Visible = False
            ElseIf pnTpoInsConv = 2 And pbSecSalud = False Then
                Label1.Visible = True
                txtTotalRemBrutaTotal.Visible = True
            End If
    Else
    
        Call CargarFlexEdit(pnTpoInsConv, pbSecSalud)
            
            nFilaPrimero2 = 0
            vnTotal2 = 0
            nTotalCompraDeu2 = 0
            
            nOrigPrimera2 = nFilaPrimero2
            nOrigTotal2 = vnTotal2
            nOrigCompraDeu2 = nTotalCompraDeu2
            
            If pnTpoInsConv = 1 And pbSecSalud = True Then
                feRemBrutaTotal.EliminaFila (2)
            End If
            
            If (pnTpoInsConv = 1 And pbSecSalud = False) Or (pnTpoInsConv = 1 And pbSecSalud = True) Then
                Label1.Visible = False
                txtTotalRemBrutaTotal.Visible = False
            ElseIf pnTpoInsConv = 2 And pbSecSalud = False Then
                Label1.Visible = True
                txtTotalRemBrutaTotal.Visible = True
            End If
            
    End If
    
    Me.Show 1
    
If IsArray(MtrRembruTotal2) Then
    pMtrRemuBrutaTotal2 = MtrRembruTotal2
    psTotalCompraDeuda = nTotalCompraDeu2
    psTotal = vnTotal2
    psFilaPrimero = nFilaPrimero2
End If
    
End Sub

Private Sub CargarGridConArray1()
    Dim i As Integer
    Dim nTotal As Double

    feRemBrutaTotal.lbEditarFlex = True
    
    For i = 1 To UBound(MtrRembruTotal1, 2)
        feRemBrutaTotal.AdicionaFila
        feRemBrutaTotal.TextMatrix(i, 1) = MtrRembruTotal1(1, i)
        feRemBrutaTotal.TextMatrix(i, 2) = MtrRembruTotal1(2, i)
        feRemBrutaTotal.TextMatrix(i, 3) = Format(MtrRembruTotal1(3, i), "#,##0.00")
        
        If feRemBrutaTotal.TextMatrix(i, 1) = "1" Then
            
            Else
                nTotal = nTotal + feRemBrutaTotal.TextMatrix(i, 3)
        End If
    Next i
    nFilaPrimero1 = feRemBrutaTotal.TextMatrix(1, 3)
    txtTotalRemBrutaTotal.Text = Format(nTotal, "#,##0.00")
    
    Select Case CInt(feRemBrutaTotal.TextMatrix(feRemBrutaTotal.row, 1))
    Case 4
         Me.feRemBrutaTotal.ForeColorRow vbBlack, True
    End Select
    
    nOrigPrimera1 = nFilaPrimero1
    nOrigTotal1 = txtTotalRemBrutaTotal.Text
    
End Sub

Private Sub CargarGridConArray2()
    Dim i As Integer
    Dim nTotal As Double

    feRemBrutaTotal.lbEditarFlex = True
    For i = 1 To UBound(MtrRembruTotal2, 2)
        feRemBrutaTotal.AdicionaFila
        feRemBrutaTotal.TextMatrix(i, 1) = MtrRembruTotal2(1, i)
        feRemBrutaTotal.TextMatrix(i, 2) = MtrRembruTotal2(2, i)
        feRemBrutaTotal.TextMatrix(i, 3) = Format(MtrRembruTotal2(3, i), "#,##0.00")
        
        If feRemBrutaTotal.TextMatrix(i, 1) = "1" Then
            
            Else
                nTotal = nTotal + feRemBrutaTotal.TextMatrix(i, 3)
        End If
    Next i
    
    nFilaPrimero2 = feRemBrutaTotal.TextMatrix(1, 3)
    txtTotalRemBrutaTotal.Text = Format(nTotal, "#,##0.00")
        
    Select Case CInt(feRemBrutaTotal.TextMatrix(feRemBrutaTotal.row, 1))
    Case 4
         Me.feRemBrutaTotal.ForeColorRow vbBlack, True
    End Select
    
    nOrigPrimera2 = nFilaPrimero2
    nOrigTotal2 = txtTotalRemBrutaTotal.Text
End Sub

'Cargar Datos en la grilla Remuneracion Bruta Total Mes 1
Private Sub CargarFlexEdit(ByVal pnTipoInst As Integer, ByVal pbSecSalud As Boolean)
    Dim lnFila As Integer
    Dim CargarFlexEdit As Boolean
    
    Dim oNCOMFormatosEval As New COMNCredito.NCOMFormatosEval
    
    'CargarFlexEdit = oNCOMFormatosEval.CargaDatosFlexEdit(8, fsCtaCod, , , , , , , , , , , , , , , , rsDatRemBrutaTotalMes1, rsDatRemBrutaTotalMes2)
    CargarFlexEdit = oNCOMFormatosEval.CargaDatosFlexEdit(8, fsCtaCod, , , , , , , , , , , , , , , , rsDatRemBrutaTotalMes1, rsDatRemBrutaTotalMes2, pnTipoInst, pbSecSalud)
   
   If nNum = 1 Then
        'Remuneracion Bruta Total Mes 1
            feRemBrutaTotal.Clear
            feRemBrutaTotal.FormaCabecera
            'feRemBrutaTotal.Rows = 2
            Call LimpiaFlex(feRemBrutaTotal)
                Do While Not rsDatRemBrutaTotalMes1.EOF
                    feRemBrutaTotal.AdicionaFila
                    lnFila = feRemBrutaTotal.row
                    feRemBrutaTotal.TextMatrix(lnFila, 1) = rsDatRemBrutaTotalMes1!nConsValor
                    feRemBrutaTotal.TextMatrix(lnFila, 2) = rsDatRemBrutaTotalMes1!cConsDescripcion
                    feRemBrutaTotal.TextMatrix(lnFila, 3) = Format(rsDatRemBrutaTotalMes1!nMonto, "#,##0.00")
                    rsDatRemBrutaTotalMes1.MoveNext
                Loop
            rsDatRemBrutaTotalMes1.Close
            Set rsDatRemBrutaTotalMes1 = Nothing
    Else
        'Remuneracion Bruta Total Mes 2
            feRemBrutaTotal.Clear
            feRemBrutaTotal.FormaCabecera
            'feRemBrutaTotal.Rows = 2
            Call LimpiaFlex(feRemBrutaTotal)
                Do While Not rsDatRemBrutaTotalMes2.EOF
                    feRemBrutaTotal.AdicionaFila
                    lnFila = feRemBrutaTotal.row
                    feRemBrutaTotal.TextMatrix(lnFila, 1) = rsDatRemBrutaTotalMes2!nConsValor
                    feRemBrutaTotal.TextMatrix(lnFila, 2) = rsDatRemBrutaTotalMes2!cConsDescripcion
                    feRemBrutaTotal.TextMatrix(lnFila, 3) = Format(rsDatRemBrutaTotalMes2!nMonto, "#,##0.00")
                    rsDatRemBrutaTotalMes2.MoveNext
                Loop
               
            rsDatRemBrutaTotalMes2.Close
            Set rsDatRemBrutaTotalMes2 = Nothing
    End If
    
End Sub
'FIN Cargar Datos en la Remuneracion Bruta Total Mes 1

Private Sub feRemBrutaTotal_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim Editar() As String
    
    Editar = Split(Me.feRemBrutaTotal.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{Tab}"
        Exit Sub
    End If
    
End Sub

Private Sub feRemBrutaTotal_RowColChange()
     If feRemBrutaTotal.Col = 3 Then
        feRemBrutaTotal.AvanceCeldas = Vertical
    Else
        feRemBrutaTotal.AvanceCeldas = Horizontal
    End If
    
    Select Case CInt(feRemBrutaTotal.TextMatrix(feRemBrutaTotal.row, 1))
      Case 4
        'Me.feRemBrutaTotal.ColumnasAEditar = "X-X-X-X"
      Case Else
        Me.feRemBrutaTotal.ColumnasAEditar = "X-X-X-3"
    End Select
         
End Sub

Public Function DisableCloseButton(frm As Form) As Boolean
'PURPOSE: Removes X button from a form
'EXAMPLE: DisableCloseButton Me
'RETURNS: True if successful, false otherwise
'NOTES:   Also removes Exit Item from
'         Control Box Menu
    Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
    
    lHndSysMenu = GetSystemMenu(frm.hwnd, 0)

    'remove close button
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)

   'Remove seperator bar
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
    
    'Return True if both calls were successful
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Select Case CInt(feRemBrutaTotal.TextMatrix(feRemBrutaTotal.row, 1))
      Case 1, 2, 3, 4, 5, 6
        Me.feRemBrutaTotal.ColumnasAEditar = "X-X-X-X"
    End Select
    End If
End Sub

Private Sub Form_Load()
DisableCloseButton Me
CentraForm Me
cmdAgregarRemBrutaTotal.Enabled = False
cmdQuitarRemBrutaTotal.Enabled = False

Label1.Visible = False
txtTotalRemBrutaTotal.Visible = False
End Sub

Private Function Calcular() As Currency
    Dim nIndice As Integer
    
    For nIndice = 2 To feRemBrutaTotal.Rows - 1
        Calcular = Calcular + feRemBrutaTotal.TextMatrix(nIndice, 3)
    Next
End Function
