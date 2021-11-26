VERSION 5.00
Begin VB.Form frmCredFormEvalReciboHonorarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibo Por Honorarios"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   Icon            =   "frmCredFormEvalReciboHonorarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
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
      Left            =   3840
      TabIndex        =   6
      Top             =   2440
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregarRecHonorarios 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitarRecHonorarios 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptarRecHonorarios 
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
      Left            =   2760
      TabIndex        =   1
      Top             =   2440
      Width           =   975
   End
   Begin VB.TextBox txtRecHonorariosPromedio 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Top             =   2100
      Width           =   1215
   End
   Begin SICMACT.FlexEdit fgReciboHonorarios 
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   3413
      Cols0           =   5
      HighLight       =   1
      EncabezadosNombres=   "N°-Año-Mes-Monto-Aux"
      EncabezadosAnchos=   "450-1300-1300-1700-0"
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
      ColumnasAEditar =   "X-1-2-3-X"
      ListaControles  =   "0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-R-C"
      FormatosEdit    =   "0-0-0-2-0"
      TextArray0      =   "N°"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   3
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   450
      RowHeight0      =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Promedio"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   2100
      Width           =   735
   End
End
Attribute VB_Name = "frmCredFormEvalReciboHonorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre      : frmCredFormEvalReciboHonorios
'** Descripción : Formulario para evaluación de Creditos            '
'** Referencia  : ERS004-2016                                       '
'** Creación    : JOEP, 20160525 09:00:00 AM                        '
'*******************************************************************'

Option Explicit
Dim Suma As Double
Dim Total As Double
Dim vnTotal As Double

Dim MtrRembruTotal As Variant
Dim i As Integer

Dim nValorOriginal As Currency


Private Sub cmdAceptarRecHonorarios_Click()

If fgReciboHonorarios.TextMatrix(1, 0) <> "" Then 'Add JOEP20210713 Mejora registrar ReciboHonorarios
    If Validar Then
        If fgReciboHonorarios.TextMatrix(1, 2) = "" Then
            MsgBox "Ud. primero debe registrar Recibo Por Honorarios", vbInformation, "Aviso"
            Exit Sub
        End If
        
        Select Case CInt(fgReciboHonorarios.TextMatrix(fgReciboHonorarios.row, 1))
        Case Is > 0
        Case Else
            MsgBox "Año mal Ingresado", vbInformation, "Alerta"
            Exit Sub
        End Select
            
        Select Case CInt(fgReciboHonorarios.TextMatrix(fgReciboHonorarios.row, 2))
        Case Is > 0
        Case Else
            MsgBox "Mes mal Ingresado, Ejemplo Enero=01", vbInformation, "Alerta"
            Exit Sub
        End Select
            
        Call Calculo
        ReDim MtrRembruTotal(3, 0)
            For i = 1 To fgReciboHonorarios.rows - 1
                ReDim Preserve MtrRembruTotal(3, i)
                MtrRembruTotal(1, i) = fgReciboHonorarios.TextMatrix(i, 1)
                MtrRembruTotal(2, i) = fgReciboHonorarios.TextMatrix(i, 2)
                MtrRembruTotal(3, i) = fgReciboHonorarios.TextMatrix(i, 3)
            Next i
        Unload Me
    End If
'Add JOEP20210713 Mejora registrar ReciboHonorarios
Else
    Set MtrRembruTotal = Nothing
    Unload Me
End If
'Add JOEP20210713 Mejora registrar ReciboHonorarios

End Sub

'Agrgar fila a la grilla
Private Sub cmdAgregarRecHonorarios_Click()
    If fgReciboHonorarios.rows <= 6 Then
            fgReciboHonorarios.lbEditarFlex = True
            fgReciboHonorarios.AdicionaFila
            fgReciboHonorarios.SetFocus
            SendKeys "{Enter}"
        Else
            MsgBox "No puede agregar mas de 7 registros", vbInformation, "Aviso"
    End If
    cmdAceptarRecHonorarios.Enabled = True
End Sub

'Quitar fila a la grilla
Private Sub cmdQuitarRecHonorarios_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        fgReciboHonorarios.EliminaFila (fgReciboHonorarios.row)
        Call Calculo
    End If
End Sub

Private Sub cmdCancelar_Click()
vnTotal = nValorOriginal
Unload Me
End Sub

'Cuando doy enter Capturo el valor
Private Sub fgReciboHonorarios_OnCellChange(pnRow As Long, pnCol As Long)
       
    Select Case pnCol
    Case 1
        If IsNumeric(fgReciboHonorarios.TextMatrix(pnRow, pnCol)) Then
            
            If Len(fgReciboHonorarios.TextMatrix(pnRow, pnCol)) = 4 Then
            Else
                MsgBox "Solo se permite 4 digitos", vbInformation, "Alerta"
                fgReciboHonorarios.TextMatrix(pnRow, pnCol) = Format("0", "00##")
                Exit Sub
            End If
            
            Select Case CCur(fgReciboHonorarios.TextMatrix(pnRow, pnCol))
                Case Is >= 0
                Case Else
                    MsgBox "Año mal Ingresado", vbInformation, "Alerta"
                    fgReciboHonorarios.TextMatrix(pnRow, pnCol) = Format("0", "00##")
                    Exit Sub
            End Select
        Else
            fgReciboHonorarios.TextMatrix(pnRow, pnCol) = Format("0", "00##")
        End If
    Case 2
        If IsNumeric(fgReciboHonorarios.TextMatrix(pnRow, pnCol)) Then
            
            If Len(fgReciboHonorarios.TextMatrix(pnRow, pnCol)) = 2 Then
            Else
                MsgBox "Solo se permite 02 digitos, Ejemplo Enero=01", vbInformation, "Alerta"
                fgReciboHonorarios.TextMatrix(pnRow, pnCol) = Format("0", "0#")
                Exit Sub
            End If
            
             Select Case CCur(fgReciboHonorarios.TextMatrix(pnRow, pnCol))
                Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
                Case Else
                    MsgBox "Mes mal Ingresado, Ejemplo Enero=01", vbInformation, "Alerta"
                    fgReciboHonorarios.TextMatrix(pnRow, pnCol) = Format("0", "0#")
                    Exit Sub
            End Select
        Else
            fgReciboHonorarios.TextMatrix(pnRow, pnCol) = Format("0", "0#")
        End If
    Case 3
        If IsNumeric(fgReciboHonorarios.TextMatrix(pnRow, pnCol)) Then
            Select Case CCur(fgReciboHonorarios.TextMatrix(pnRow, pnCol))
                Case Is >= 0
                Case Else
                    MsgBox "Monto mal Ingresado", vbInformation, "Alerta"
                    fgReciboHonorarios.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
                    Exit Sub
            End Select
            fgReciboHonorarios.TextMatrix(pnRow, pnCol) = UCase(fgReciboHonorarios.TextMatrix(pnRow, pnCol))
            Call Calculo
        Else
            MsgBox "Monto mal Ingresado", vbInformation, "Alerta"
                    fgReciboHonorarios.TextMatrix(pnRow, pnCol) = Format("0", "#,##0.00")
        End If
    End Select
End Sub

'Para Calcular El Monto Promedio
Public Sub Calculo()
Dim cont As Integer
    
    txtRecHonorariosPromedio.Text = Format(SumarCampo(fgReciboHonorarios, 3), "#,##0.00")
    Suma = txtRecHonorariosPromedio.Text
    cont = fgReciboHonorarios.rows - 1
      
    Total = Suma / cont
    
    txtRecHonorariosPromedio.Text = Format(Total, "#,##0.00")
    vnTotal = txtRecHonorariosPromedio.Text
End Sub

Public Sub Inicio(ByRef psTotal As Double, ByRef pMtrReciboHonorarios As Variant)

If IsArray(pMtrReciboHonorarios) Then
        MtrRembruTotal = pMtrReciboHonorarios
        Call CargarGridConArray
    Else
    txtRecHonorariosPromedio.Text = "0.00"
        vnTotal = 0
End If

Me.Show 1

'Comento JOEP20210713 Mejora registrar ReciboHonorarios
'psTotal = vnTotal
'pMtrReciboHonorarios = MtrRembruTotal
'Comento JOEP20210713 Mejora registrar ReciboHonorarios

'Add JOEP20210713 Mejora registrar ReciboHonorarios
If IsArray(MtrRembruTotal) Then
    psTotal = vnTotal
    pMtrReciboHonorarios = MtrRembruTotal
Else
    psTotal = 0
    Set pMtrReciboHonorarios = Nothing
End If
'Add JOEP20210713 Mejora registrar ReciboHonorarios
End Sub

Private Sub CargarGridConArray()
    Dim i As Integer
    Dim cont As Integer
    Dim nSuma As Double
    Dim nTotal As Double

    fgReciboHonorarios.lbEditarFlex = True
    For i = 1 To UBound(MtrRembruTotal, 2)
        fgReciboHonorarios.AdicionaFila
        fgReciboHonorarios.TextMatrix(i, 1) = MtrRembruTotal(1, i)
        fgReciboHonorarios.TextMatrix(i, 2) = Format(MtrRembruTotal(2, i), "0#")
        fgReciboHonorarios.TextMatrix(i, 3) = MtrRembruTotal(3, i)
        nSuma = nSuma + fgReciboHonorarios.TextMatrix(i, 3)
    Next i
    cont = fgReciboHonorarios.rows - 1
        
        nTotal = nSuma / cont
    txtRecHonorariosPromedio.Text = Format(nTotal, "#,##0.00")
    
    nValorOriginal = txtRecHonorariosPromedio.Text
End Sub

Private Sub Form_Load()
CentraForm Me
'Comento JOEP20210713 Mejora registrar ReciboHonorarios
'cmdAceptarRecHonorarios.Enabled = False
'Comento JOEP20210713 Mejora registrar ReciboHonorarios

'Add JOEP20210713 Mejora registrar ReciboHonorarios
cmdAceptarRecHonorarios.Enabled = True
'Add JOEP20210713 Mejora registrar ReciboHonorarios
End Sub

Private Function Validar() As Boolean
Dim i As Integer
Validar = True

    For i = 1 To fgReciboHonorarios.rows - 1
        If fgReciboHonorarios.TextMatrix(i, 1) = "" Then
            MsgBox "Año Vacio", vbInformation, "Alerta"
            Validar = False
            Exit Function
        ElseIf fgReciboHonorarios.TextMatrix(i, 1) = 0 Then
            MsgBox "Año Mal Ingresado", vbInformation, "Alerta"
            Validar = False
            Exit Function
        ElseIf fgReciboHonorarios.TextMatrix(i, 2) = "" Then
            MsgBox "Mes Vacio", vbInformation, "Alerta"
            Validar = False
            Exit Function
        ElseIf fgReciboHonorarios.TextMatrix(i, 2) = 0 Then
            MsgBox "Mes Mal Ingresado, Ejemplo Enero=01", vbInformation, "Alerta"
            Validar = False
            Exit Function
        ElseIf fgReciboHonorarios.TextMatrix(i, 3) = "" Then
            MsgBox "Monto Vacio", vbInformation, "Alerta"
            Validar = False
            Exit Function
        End If
    Next i
End Function

