VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRHCalculoEPS 
   Caption         =   "Calculo EPS"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12195
   Icon            =   "frmRHCalculoEPS.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   12195
   Begin VB.CommandButton cmdArchivar 
      Caption         =   "Archivar"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   6165
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   6120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10800
      TabIndex        =   1
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Export  Excel >>"
      Height          =   375
      Left            =   9480
      TabIndex        =   2
      ToolTipText     =   "Exportar a la Plantilla Mensual"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdMensual 
      Caption         =   "Export Mens >>"
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      ToolTipText     =   "Exportar a la Plantilla Mensual"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuincena 
      Caption         =   "Export Quin >>"
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      ToolTipText     =   "Exportar A la Plantilla de la Quincena"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAjustar 
      Caption         =   "Ajustar Ventana"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFEPS 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   9763
      _Version        =   393216
      FixedCols       =   0
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Sicmact.TxtBuscar txtCodCalculo 
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   2100
      _extentx        =   3704
      _extenty        =   503
      appearance      =   1
      appearance      =   1
      font            =   "frmRHCalculoEPS.frx":030A
      appearance      =   1
      tipobusqueda    =   2
      stitulo         =   ""
   End
   Begin VB.Label lblfecha 
      Caption         =   "fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9840
      TabIndex        =   12
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2100
      TabIndex        =   8
      Top             =   120
      Width           =   4740
   End
   Begin VB.OLE OleExcel 
      Class           =   "Excel.Sheet.8"
      Height          =   255
      Left            =   1680
      OleObjectBlob   =   "frmRHCalculoEPS.frx":0336
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmRHCalculoEPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oAsisMedica As DActualizaAsistMedicaPrivada
Dim rs As New ADODB.Recordset
Dim bAjuste As Boolean
Dim Progress As clsProgressBar

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim bPermiso As Boolean

Private Sub cmdAjustar_Click()
 If bAjuste = True Then
    MSHFEPS.ColWidth(3) = 850
    MSHFEPS.ColWidth(4) = 750
    MSHFEPS.ColWidth(13) = 750
    bAjuste = False
  Else
    MSHFEPS.ColWidth(3) = 0
    MSHFEPS.ColWidth(4) = 0
    MSHFEPS.ColWidth(13) = 0
    MSHFEPS.Refresh
    bAjuste = True
  End If
End Sub

Private Sub cmdArchivar_Click()
Me.cmdArchivar.Visible = False
Me.cmdGrabar.Visible = True
Me.txtDescripcion.Visible = True

End Sub

Private Sub cmdExportar_Click()
 Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    If Me.MSHFEPS.TextMatrix(1, 1) = "" Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Date), "yyyy") & Format(Time, "hhmmss") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       GeneraReporteEPS MSHFEPS, xlHoja1
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
       OleExcel.Appearance = 0
       OleExcel.Width = 500
    End If
    MousePointer = 0
End Sub

Private Sub cmdGrabar_Click()
Dim i As Long
Dim J As Long
Dim oCon As NRHConcepto
Set oCon = New NRHConcepto
Dim sCodPersona As String
Dim nMonto As String
Dim sCodConceptoEPS As String
Me.cmdArchivar.Visible = True
Me.cmdGrabar.Visible = False
Me.txtDescripcion.Visible = False

If MsgBox("¿Desea Grabar Los datos  del Calculo de EPS ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
'Cabecera
'AgregaCalculoEPS(psCodCalculoEps As String, psDescripcion As String, psUltimaActualziacion As String, pnEstadoEPS As Integer) As Integer
sCodConceptoEPS = GetMovNro(gsCodUser, gsCodAge)

If oAsisMedica.GetRHExistCodEPS(Format(gdFecSis, "YYYYMMDD")) > 1 Then
    If MsgBox("Ya existen Registros del dia " + Format(gdFecSis, "YYYY/MM/DD") + " Desea Continuar ? ", vbYesNo + vbQuestion, " Verificando Grabacion") = vbNo Then Exit Sub
End If
oPlaEvento_ShowProgress
oCon.AgregaCalculoEPS sCodConceptoEPS, txtDescripcion.Text, sCodConceptoEPS, 0

For i = 1 To MSHFEPS.Rows - 2
    sCodPersona = MSHFEPS.TextMatrix(i, 0)
    'psCodCalculoEps, psCodPersona, pnSueldo, pnSueldo_x_225, pnCantPersonas, pnPlanSinIGV, pnPromedio, _
     pnNeto, pnPagaEmpleado, pnPagaEmpresa, pnAdicionalHijos, pnAdicionalPadres, pnTotalEmpleado, _
     pnDesQuincena, pnDesMensual, pnDesQuincenaUno, pnDescQuincenaDos, pnSaldo
   
    oCon.AgregaRHCalculoEPSDet sCodConceptoEPS, MSHFEPS.TextMatrix(i, 0), MSHFEPS.TextMatrix(i, 2), MSHFEPS.TextMatrix(i, 3), _
    MSHFEPS.TextMatrix(i, 4), MSHFEPS.TextMatrix(i, 5), MSHFEPS.TextMatrix(i, 6), MSHFEPS.TextMatrix(i, 7), MSHFEPS.TextMatrix(i, 8), _
    MSHFEPS.TextMatrix(i, 9), MSHFEPS.TextMatrix(i, 10), MSHFEPS.TextMatrix(i, 11), MSHFEPS.TextMatrix(i, 12), MSHFEPS.TextMatrix(i, 13), _
    MSHFEPS.TextMatrix(i, 14), MSHFEPS.TextMatrix(i, 15), MSHFEPS.TextMatrix(i, 16), MSHFEPS.TextMatrix(i, 17)
    
    oPlaEvento_Progress i, MSHFEPS.Rows - 1
Next
    oCon.ActualizaRHCalEPSDet sCodConceptoEPS

oPlaEvento_CloseProgress

End Sub

Private Sub cmdMensual_Click()
Dim i As Long
Dim J As Long
Dim oCon As NRHConcepto
Set oCon = New NRHConcepto
Dim sCodPersona As String
Dim nMonto As String
If MsgBox("¿ Desea Grabar los  datos en la Planilla Mensual?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
oPlaEvento_ShowProgress
For i = 1 To MSHFEPS.Rows - 2
    sCodPersona = MSHFEPS.TextMatrix(i, 0)
    nMonto = MSHFEPS.TextMatrix(i, 14)
    If nMonto < 0 Then
       nMonto = 0
    End If
    oCon.EliminaConceptoRRHH sCodPersona, "E01", "228"
    oCon.AgregaConceptoRRHH sCodPersona, "E01", "228", nMonto, GetMovNro(gsCodUser, gsCodAge)
    oPlaEvento_Progress i, MSHFEPS.Rows - 1
Next
oPlaEvento_CloseProgress

End Sub

Private Sub cmdQuincena_Click()
Dim i As Long
Dim J As Long
Dim oCon As NRHConcepto
Set oCon = New NRHConcepto
Dim sCodPersona As String
Dim nMonto As String


If MsgBox("¿ Desea Grabar los  datos en la Planilla de Quincena ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub

oPlaEvento_ShowProgress
For i = 1 To MSHFEPS.Rows - 2
    sCodPersona = MSHFEPS.TextMatrix(i, 0)
    nMonto = MSHFEPS.TextMatrix(i, 13)
    oCon.EliminaConceptoRRHH sCodPersona, "E21", "228"
    oCon.AgregaConceptoRRHH sCodPersona, "E21", "228", nMonto, GetMovNro(gsCodUser, gsCodAge)
    oPlaEvento_Progress i, MSHFEPS.Rows - 1
Next
oPlaEvento_CloseProgress



End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblfecha.Caption = gdFecSis

Me.Width = 12195
Me.Height = 7020
Set oAsisMedica = New DActualizaAsistMedicaPrivada
Set rs = New ADODB.Recordset
Set Progress = New clsProgressBar
bAjuste = True
'Set rs = oAsisMedica.GetRHCalculoEPS(20)
'Set MSHFEPS.DataSource = rs
bPermiso = False
txtCodCalculo.rs = oAsisMedica.GetRHCodcalculoEPS
txtCodCalculo.Text = ""
bPermiso = True

Calcular_eps frmRHAsignacionPlan.txtPromedioEPS.Text, Trim(frmRHAsignacionPlan.txtano) + Trim(frmRHAsignacionPlan.txtMes), frmRHAsignacionPlan.meQuincena, frmRHAsignacionPlan.meEmpresa

MSHFEPS.ColWidth(0) = 0
MSHFEPS.ColWidth(1) = 1500
MSHFEPS.ColWidth(2) = 700
MSHFEPS.ColWidth(3) = 0
MSHFEPS.ColWidth(4) = 0
MSHFEPS.ColWidth(5) = 760
MSHFEPS.ColWidth(6) = 760
MSHFEPS.ColWidth(7) = 760
MSHFEPS.ColWidth(8) = 760
MSHFEPS.ColWidth(9) = 760
MSHFEPS.ColWidth(10) = 760
MSHFEPS.ColWidth(11) = 760
MSHFEPS.ColWidth(12) = 760
MSHFEPS.ColWidth(13) = 760

MSHFEPS.ColAlignmentFixed(2) = 3
MSHFEPS.ColAlignmentFixed(3) = 3
MSHFEPS.ColAlignmentFixed(4) = 3
MSHFEPS.ColAlignmentFixed(5) = 3
MSHFEPS.ColAlignmentFixed(6) = 3
MSHFEPS.ColAlignmentFixed(7) = 3
MSHFEPS.ColAlignmentFixed(8) = 3
MSHFEPS.ColAlignmentFixed(9) = 3
MSHFEPS.ColAlignmentFixed(10) = 3
MSHFEPS.ColAlignmentFixed(11) = 3
MSHFEPS.ColAlignmentFixed(12) = 3
MSHFEPS.ColAlignmentFixed(13) = 3
MSHFEPS.ColAlignmentFixed(14) = 3

MSHFEPS.ColAlignment(2) = 6
MSHFEPS.ColAlignment(3) = 6
MSHFEPS.ColAlignment(4) = 3
MSHFEPS.ColAlignment(5) = 6
MSHFEPS.ColAlignment(6) = 6
MSHFEPS.ColAlignment(7) = 6
MSHFEPS.ColAlignment(8) = 6
MSHFEPS.ColAlignment(9) = 6
MSHFEPS.ColAlignment(10) = 6
MSHFEPS.ColAlignment(11) = 6
MSHFEPS.ColAlignment(12) = 6
MSHFEPS.ColAlignment(13) = 6
MSHFEPS.ColAlignment(14) = 6

Habilitar_Botones True

End Sub


Sub Calcular_eps(pnPromedio As Currency, psPeriodo As String, pnQuincena As Currency, pnEmpresa As Currency)
Set rs = oAsisMedica.GetRHCalculoEPS(pnPromedio, psPeriodo, pnQuincena, pnEmpresa)
Set MSHFEPS.DataSource = rs
If MSHFEPS.TextMatrix(MSHFEPS.Rows - 1, 2) <> "" Then
        MSHFEPS.Rows = MSHFEPS.Rows + 1
        MSHFEPS.TextMatrix(MSHFEPS.Rows - 1, 3) = "TOTAL"
        For J = 2 To Me.MSHFEPS.Cols - 1
            lnAcumulador = 0
            If Left(MSHFEPS.TextMatrix(0, J), 2) <> "U_" And Left(MSHFEPS.TextMatrix(0, J), 1) <> "_" Then
                For i = 1 To Me.MSHFEPS.Rows - 2
                    If MSHFEPS.TextMatrix(i, J) <> "" Then
                        If J = 14 And CCur(MSHFEPS.TextMatrix(i, J)) < 0 Then
                            MSHFEPS.TextMatrix(i, J) = 0
                        End If
                        lnAcumulador = lnAcumulador + CCur(MSHFEPS.TextMatrix(i, J))
                            Select Case J
                                Case 8, 10, 11, 12
                                MSHFEPS.Col = J
                                MSHFEPS.Row = i
                                MSHFEPS.CellBackColor = RGB(100, 200, 350)
                                Case 5, 6, 7
                                MSHFEPS.Col = J
                                MSHFEPS.Row = i
                                MSHFEPS.CellBackColor = RGB(100, 200, 200)
                            End Select
                    End If
                Next i
                MSHFEPS.TextMatrix(MSHFEPS.Rows - 1, J) = Format(lnAcumulador, "#,##.00")
                MSHFEPS.Row = MSHFEPS.Rows - 1
                MSHFEPS.Col = J
                MSHFEPS.CellBackColor = &HA0C000
                'FlexPrePla.CellFontBold = True
                lnAcumulador = lnAcumulador + CCur(MSHFEPS.TextMatrix(i, J))
            End If
        Next J
    End If
MSHFEPS.TextMatrix(MSHFEPS.Rows - 1, 1) = MSHFEPS.Rows - 2
End Sub

Private Sub GeneraReporteEPS(pflex As MSHFlexGrid, pxlHoja1 As Excel.Worksheet, Optional pnColFiltroVacia As Integer = 0)
    Dim i As Integer
    Dim K As Integer
    Dim J As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    
    For i = 0 To pflex.Rows - 1
        If pnColFiltroVacia = 0 Then
            For J = 0 To pflex.Cols - 1
                pxlHoja1.Cells(i + 1, J + 1) = pflex.TextMatrix(i, J)
            Next J
        Else
            If pflex.TextMatrix(i, pnColFiltroVacia) <> "" Then
                For J = 0 To pflex.Cols - 1
                    pxlHoja1.Cells(i + 1, J + 1) = pflex.TextMatrix(i, J)
                Next J
            End If
        End If
    Next i
    
End Sub


Private Sub oPlaEvento_ShowProgress()
    Progress.ShowForm Me
End Sub

Private Sub oPlaEvento_Progress(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Actualizando Descuento EPS Planilla Quincenal"
End Sub
Private Sub oPlaEvento_Progress2(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Actualizando Descuento EPS Planilla Mensual"
End Sub

Private Sub oPlaEvento_CloseProgress()
    Progress.CloseForm Me
End Sub

Private Sub txtCodCalculo_EmiteDatos()
If bPermiso = False Then Exit Sub

lblDescripcion.Caption = txtCodCalculo.psDescripcion

Set rs = oAsisMedica.GetRHDetallecalculoEPS(txtCodCalculo.Text)
Set MSHFEPS.Recordset = rs
If rs.EOF = True Then
    MsgBox "No existen registros", vbInformation, "No existen Registros"
    Exit Sub
End If

If MSHFEPS.TextMatrix(MSHFEPS.Rows - 1, 2) <> "" Then
        MSHFEPS.Rows = MSHFEPS.Rows + 1
        MSHFEPS.TextMatrix(MSHFEPS.Rows - 1, 3) = "TOTAL"
        For J = 2 To Me.MSHFEPS.Cols - 1
            lnAcumulador = 0
            If Left(MSHFEPS.TextMatrix(0, J), 2) <> "U_" And Left(MSHFEPS.TextMatrix(0, J), 1) <> "_" Then
                For i = 1 To Me.MSHFEPS.Rows - 2
                    If MSHFEPS.TextMatrix(i, J) <> "" Then
                        If J = 14 And CCur(MSHFEPS.TextMatrix(i, J)) < 0 Then
                            MSHFEPS.TextMatrix(i, J) = 0
                        End If
                        lnAcumulador = lnAcumulador + CCur(MSHFEPS.TextMatrix(i, J))
                            Select Case J
                                Case 8, 10, 11, 12
                                MSHFEPS.Col = J
                                MSHFEPS.Row = i
                                MSHFEPS.CellBackColor = RGB(300, 150, 150)
                                'RGB(100, 200, 350)
                                Case 5, 6, 7
                                MSHFEPS.Col = J
                                MSHFEPS.Row = i
                                MSHFEPS.CellBackColor = RGB(100, 200, 200)
                            End Select
                    End If
                Next i
                MSHFEPS.TextMatrix(MSHFEPS.Rows - 1, J) = Format(lnAcumulador, "#,##.00")
                MSHFEPS.Row = MSHFEPS.Rows - 1
                MSHFEPS.Col = J
                MSHFEPS.CellBackColor = &HA0C000
                'FlexPrePla.CellFontBold = True
                lnAcumulador = lnAcumulador + CCur(MSHFEPS.TextMatrix(i, J))
            End If
        Next J
    End If
MSHFEPS.TextMatrix(MSHFEPS.Rows - 1, 1) = MSHFEPS.Rows - 2

Habilitar_Botones False

End Sub

Sub Habilitar_Botones(pbvalor As Boolean)
cmdArchivar.Enabled = pbvalor
txtDescripcion.Enabled = pbvalor
cmdGrabar.Enabled = pbvalor
cmdAjustar.Enabled = pbvalor
cmdQuincena.Enabled = pbvalor
cmdMensual.Enabled = pbvalor
cmdExportar.Enabled = pbvalor
End Sub



