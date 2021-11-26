VERSION 5.00
Begin VB.Form frmCredFormEvalBoletaPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boletas de Pago"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5070
   Icon            =   "frmCredFormEvalBoletaPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin SICMACT.EditMoney edMontoNeto 
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
   End
   Begin SICMACT.EditMoney edMontoBruto 
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
   End
   Begin VB.CommandButton cmdCancelar 
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
      Left            =   3960
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptarBolPago 
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
      Left            =   2880
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitarBolPago 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregarBolPago 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin SICMACT.FlexEdit fgBoletaPago 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   2778
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "N°-Año-Mes-Monto Bruto-Monto Neto-Aux"
      EncabezadosAnchos=   "400-900-900-1200-1200-0"
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
      ColumnasAEditar =   "X-1-2-3-4-X"
      ListaControles  =   "0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-R-R-R-R-L"
      FormatosEdit    =   "0-3-3-2-2-0"
      CantEntero      =   15
      TextArray0      =   "N°"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      TipoBusPersona  =   2
   End
   Begin VB.Label Label2 
      Caption         =   "Promedio Bruto:"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Promedio Neto:"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "frmCredFormEvalBoletaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre      : frmCredFormEvalBoletaPago                         '
'** Descripción : Formulario para evaluación de Creditos            '
'** Referencia  : ERS004-2016                                       '
'** Creación    : JOEP, 20160525 09:00:00 AM                        '
'*******************************************************************'
Option Explicit
Dim Suma As Double
Dim Total As Double
Dim vnTotal As Double
Dim MtrRembruTotal As Variant

Dim nValorOriginal As Currency

Private Sub cmdCancelar_Click()
    vnTotal = nValorOriginal
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    cmdAceptarBolPago.Enabled = False
End Sub

Private Sub cmdAceptarBolPago_Click()
Dim i As Integer
If Validar Then
    Call CalculoNeto
    ReDim MtrRembruTotal(4, 0)
        For i = 1 To (fgBoletaPago.rows - 1)
            ReDim Preserve MtrRembruTotal(4, i)
            MtrRembruTotal(1, i) = fgBoletaPago.TextMatrix(i, 1)
            MtrRembruTotal(2, i) = fgBoletaPago.TextMatrix(i, 2)
            MtrRembruTotal(3, i) = fgBoletaPago.TextMatrix(i, 3)
            MtrRembruTotal(4, i) = fgBoletaPago.TextMatrix(i, 4) 'ACTA Nº 112-2018 JOEP20180614
        Next i
 Unload Me
End If
End Sub

'Agregar Fila a la grilla
Private Sub cmdAgregarBolPago_Click()
    If fgBoletaPago.rows <= 4 Then
            fgBoletaPago.lbEditarFlex = True
            fgBoletaPago.AdicionaFila
            fgBoletaPago.SetFocus
            SendKeys "{Enter}"
        Else
            MsgBox "No puede agregar mas de 5 registros", vbInformation, "Aviso"
    End If
    cmdAceptarBolPago.Enabled = True
End Sub
'quitar fila de grilla
Private Sub cmdQuitarBolPago_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        fgBoletaPago.EliminaFila (fgBoletaPago.row)
        Call CalculoNeto
        edMontoBruto = Format((SumarCampo(fgBoletaPago, 3) / (fgBoletaPago.rows - 1)), "#,##0.00") 'ACTA Nº 112-2018 JOEP20180614
    End If
End Sub

'Cuando doy enter Capturo el valor
Private Sub fgBoletaPago_OnCellChange(pnRow As Long, pnCol As Long)
    Select Case pnCol
    Case 1
        If IsNumeric(fgBoletaPago.TextMatrix(pnRow, pnCol)) Then
            If Len(Replace(fgBoletaPago.TextMatrix(pnRow, pnCol), ",", "")) <> 4 Then
                MsgBox "Solo se permite 4 digitos", vbInformation, "Alerta"
                fgBoletaPago.TextMatrix(pnRow, pnCol) = Format("0", "00##")
                Exit Sub
            End If
            Select Case CCur(fgBoletaPago.TextMatrix(pnRow, pnCol))
                Case Is > 0
                Case Else
                    MsgBox "Año mal Ingresado", vbInformation, "Alerta"
                    fgBoletaPago.TextMatrix(pnRow, pnCol) = Format("0", "00##")
                    Exit Sub
            End Select
            fgBoletaPago.TextMatrix(pnRow, pnCol) = Replace(fgBoletaPago.TextMatrix(pnRow, pnCol), ",", "")
        Else
            fgBoletaPago.TextMatrix(pnRow, pnCol) = Format("0", "00##")
        End If
    Case 2
        If IsNumeric(fgBoletaPago.TextMatrix(pnRow, pnCol)) Then
            If Len(fgBoletaPago.TextMatrix(pnRow, pnCol)) = 1 Then
                fgBoletaPago.TextMatrix(pnRow, pnCol) = "0" + fgBoletaPago.TextMatrix(pnRow, pnCol)
            End If
            If Len(fgBoletaPago.TextMatrix(pnRow, pnCol)) <> 2 Then
                MsgBox "Solo se permite 02 digitos, Ejemplo Enero=01", vbInformation, "Alerta"
                fgBoletaPago.TextMatrix(pnRow, pnCol) = Format("0", "0#")
                Exit Sub
            End If
            Select Case CCur(fgBoletaPago.TextMatrix(pnRow, pnCol))
                Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
                Case Else
                    MsgBox "Mes mal Ingresado, Ejemplo Enero=01", vbInformation, "Alerta"
                    fgBoletaPago.TextMatrix(pnRow, pnCol) = Format("0", "0#")
                    Exit Sub
            End Select
        Else
            fgBoletaPago.TextMatrix(pnRow, pnCol) = Format("0", "0#")
        End If
    Case 3 ''ACTA Nº 112-2018 JOEP20180614
         If IsNumeric(fgBoletaPago.TextMatrix(pnRow, pnCol)) Then
            fgBoletaPago.TextMatrix(pnRow, 4) = Format(IIf(fgBoletaPago.TextMatrix(pnRow, 4) = "", 0, fgBoletaPago.TextMatrix(pnRow, 4)), "#,##0.00")
            If CCur(fgBoletaPago.TextMatrix(pnRow, 4)) > CCur(fgBoletaPago.TextMatrix(pnRow, pnCol)) Then
                MsgBox "N° " & CCur(fgBoletaPago.TextMatrix(pnRow, 0)) & " - El [Monto Bruto] tiene que ser mayor al [Monto Neto]", vbInformation, "Alerta"
                fgBoletaPago.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
                fgBoletaPago.TextMatrix(pnRow, 4) = Format("0", "#,##0.00")
                Call CalculoNeto
                Exit Sub
            End If
            
            Select Case CCur(fgBoletaPago.TextMatrix(pnRow, pnCol))
                Case Is >= 0
                    edMontoBruto = Format((SumarCampo(fgBoletaPago, 3) / (fgBoletaPago.rows - 1)), "#,##0.00")
                Case Else
                    MsgBox "Monto mal Ingresado", vbInformation, "Alerta"
                    fgBoletaPago.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
                    Exit Sub
            End Select
        Else
            MsgBox "Monto Bruto mal Ingresado", vbInformation, "Alerta"
            fgBoletaPago.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
        End If
     Case 4 ''ACTA Nº 112-2018 JOEP20180614
         If IsNumeric(fgBoletaPago.TextMatrix(pnRow, pnCol)) Then
            fgBoletaPago.TextMatrix(pnRow, 3) = Format(IIf(fgBoletaPago.TextMatrix(pnRow, 3) = "", 0, fgBoletaPago.TextMatrix(pnRow, 3)), "#,##0.00")
            If CCur(fgBoletaPago.TextMatrix(pnRow, pnCol)) > CCur(fgBoletaPago.TextMatrix(pnRow, 3)) Then
                MsgBox "N° " & CCur(fgBoletaPago.TextMatrix(pnRow, 0)) & " - El [Monto Bruto] tiene que ser mayor al [Monto Neto]", vbInformation, "Alerta"
                fgBoletaPago.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
                Call CalculoNeto
                Exit Sub
            End If
            
            Select Case CCur(fgBoletaPago.TextMatrix(pnRow, pnCol))
                Case Is >= 0
                    Call CalculoNeto
                Case Else
                    MsgBox "Monto Neto mal Ingresado", vbInformation, "Alerta"
                    fgBoletaPago.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
                    Exit Sub
            End Select
        Else
            MsgBox "Monto mal Ingresado", vbInformation, "Alerta"
                    fgBoletaPago.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
        End If
    End Select
End Sub

'Para Calcular El Monto Promedio
Public Sub CalculoNeto()
    edMontoNeto = Format(SumarCampo(fgBoletaPago, 4), "#,##0.00") 'se cambio la columna 3 a 4 'ACTA Nº 112-2018 JOEP20180614
    Suma = edMontoNeto
    Total = Suma / (fgBoletaPago.rows - 1)
    edMontoNeto = Format(Total, "#,##0.00")
    vnTotal = edMontoNeto
End Sub
  
Public Sub Inicio(ByRef psTotal As Double, ByRef pMtrRemuBrutaTotal As Variant)
fgBoletaPago.Clear
FormateaFlex fgBoletaPago

    If IsArray(pMtrRemuBrutaTotal) Then
        MtrRembruTotal = pMtrRemuBrutaTotal
        Call CargarGridConArray
    Else
        edMontoNeto = 0
        nValorOriginal = 0
        vnTotal = 0
        Suma = 0
        Total = 0
    End If
Me.Show 1

If IsArray(MtrRembruTotal) Then
    psTotal = vnTotal
    pMtrRemuBrutaTotal = MtrRembruTotal
End If

End Sub

Private Sub CargarGridConArray()
    Dim i As Integer
fgBoletaPago.lbEditarFlex = True
    For i = 1 To UBound(MtrRembruTotal, 2)
        fgBoletaPago.AdicionaFila
        fgBoletaPago.TextMatrix(i, 1) = MtrRembruTotal(1, i)
        fgBoletaPago.TextMatrix(i, 2) = Format(MtrRembruTotal(2, i), "0#")
        fgBoletaPago.TextMatrix(i, 3) = MtrRembruTotal(3, i)
        fgBoletaPago.TextMatrix(i, 4) = MtrRembruTotal(4, i) 'ACTA Nº 112-2018 JOEP20180614
    Next i
       
    edMontoNeto = Format((SumarCampo(fgBoletaPago, 4) / (fgBoletaPago.rows - 1)), "#,##0.00") 'se cambio la columna 3 a 4 'ACTA Nº 112-2018 JOEP20180614
    nValorOriginal = Format((SumarCampo(fgBoletaPago, 4) / (fgBoletaPago.rows - 1)), "#,##0.00") 'se cambio la columna 3 a 4 'ACTA Nº 112-2018 JOEP20180614
    
    edMontoBruto = Format((SumarCampo(fgBoletaPago, 3) / (fgBoletaPago.rows - 1)), "#,##0.00") 'ACTA Nº 112-2018 JOEP20180614
End Sub

Private Function Validar() As Boolean
Dim i As Integer
Validar = True

If fgBoletaPago.TextMatrix(1, 1) = "" Then
    MsgBox "Ud. primero debe registrar Boleta de Pago", vbInformation, "Aviso"
    Validar = False
    Exit Function
End If
    
For i = 1 To (fgBoletaPago.rows - 1)
        If fgBoletaPago.TextMatrix(i, 1) = "" Then
            MsgBox "Año Vacio", vbInformation, "Alerta"
            Validar = False
            Exit Function
        ElseIf fgBoletaPago.TextMatrix(i, 1) <= 0 Then
            MsgBox "Año Mal Ingresado", vbInformation, "Alerta"
            Validar = False
            Exit Function
        ElseIf fgBoletaPago.TextMatrix(i, 2) = "" Then
            MsgBox "Mes Vacio, Ejemplo Enero=01", vbInformation, "Alerta"
            Validar = False
            Exit Function
        ElseIf fgBoletaPago.TextMatrix(i, 2) <= 0 Then
            MsgBox "Solo se permite 02 digitos, Ejemplo Enero=01", vbInformation, "Alerta"
            Validar = False
            Exit Function
        ElseIf fgBoletaPago.TextMatrix(i, 3) = "" Then
            MsgBox "Ingrese Monto Bruto", vbInformation, "Alerta"
            Validar = False
            Exit Function
        ElseIf fgBoletaPago.TextMatrix(i, 3) <= 0 Then
            MsgBox "El Monto Bruto no puede ser 0 " & "item:" & i, vbInformation, "Alerta"
            Validar = False
            Exit Function
        ElseIf fgBoletaPago.TextMatrix(i, 4) = "" Then 'ACTA Nº 112-2018 JOEP20180614
            MsgBox "Ingrese el Monto Neto", vbInformation, "Alerta"
            Validar = False
            Exit Function
        End If
    Next i

If edMontoBruto = 0 Then
    MsgBox "El Promedio del Monto Bruto no puede ser 0", vbInformation, "Alerta"
    Validar = False
    Exit Function
End If

End Function

