VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRHReportes 
   Caption         =   "Reportes"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   Icon            =   "frmRHreportes.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11865
   Begin VB.Frame ffechas 
      Caption         =   "Fecha"
      Height          =   615
      Left            =   840
      TabIndex        =   7
      Top             =   480
      Width           =   5535
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   300
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFecFin 
         Height          =   300
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblFecINi 
         Caption         =   "Fecha Ini:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   270
         Width           =   735
      End
      Begin VB.Label lblFecFin 
         Caption         =   "Fecha Fin:"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   8520
      TabIndex        =   5
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10440
      TabIndex        =   4
      Top             =   7320
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLista 
      Height          =   6015
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10610
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox cmbReportes 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
   Begin VB.OLE OleExcel 
      Class           =   "Excel.Sheet.8"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Reporte"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "frmRHReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oReporte As DRHReportes
Dim rs As New ADODB.Recordset


Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim Progress As clsProgressBar







Private Sub cmbReportes_Click()

Select Case Right(cmbReportes.Text, 1)
Case 1
        ffechas.Visible = False

Case 2
        ffechas.Visible = False

Case 3
        ffechas.Visible = True

End Select



End Sub

Private Sub cmdExportar_Click()
Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    If Me.MSHFLista.TextMatrix(1, 1) = "" Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Date), "yyyy") & Format(Time, "hhmmss") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       GeneraReporteRH MSHFLista, xlHoja1
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

Private Sub cmdProcesar_Click()
If cmbReportes.Text = "" Then Exit Sub

Select Case Right(cmbReportes.Text, 1)
Case 1
        MSHFLista.Clear
        MSHFLista.Rows = 2
        Set rs = oReporte.GetRHlistaTrabDep
        MSHFLista.ColWidth(0) = 1300
        MSHFLista.ColWidth(1) = 3000
        MSHFLista.ColWidth(2) = 2000
        MSHFLista.ColWidth(3) = 2500
        MSHFLista.ColWidth(4) = 2500
        MSHFLista.ColWidth(5) = 2000 'fecha
        MSHFLista.ColWidth(6) = 2000 'fecha
        Set MSHFLista.DataSource = rs
Case 2
        MSHFLista.Clear
        MSHFLista.Rows = 2
        Set rs = oReporte.SP_RHlistaAnalistasAg
        MSHFLista.ColWidth(0) = 1300
        MSHFLista.ColWidth(1) = 3000
        MSHFLista.ColWidth(2) = 2000
        MSHFLista.ColWidth(3) = 2500
        MSHFLista.ColWidth(4) = 2500
        MSHFLista.ColWidth(5) = 1 'fecha
        MSHFLista.ColWidth(6) = 2000 'fecha
        Set MSHFLista.DataSource = rs
        
Case 3
        MSHFLista.Clear
        MSHFLista.Rows = 2
        MSHFLista.Cols = 11
        MSHFLista.ColWidth(0) = 500
        MSHFLista.ColWidth(1) = 980
        MSHFLista.ColWidth(2) = 2000
        MSHFLista.ColWidth(3) = 980
        MSHFLista.ColWidth(4) = 980
        MSHFLista.ColWidth(5) = 980
        MSHFLista.ColWidth(6) = 980
        MSHFLista.ColWidth(7) = 980
        MSHFLista.ColWidth(8) = 980
        MSHFLista.ColWidth(9) = 980
        MSHFLista.ColWidth(10) = 980
        ReporteQuinta gdFecSis, Me.mskFecIni.Text, Me.mskFecFin.Text, gsEmpresa
        'lsCad = oRepEvento.Rep5taRRHH(gdFecSis, Me.mskFecIni.Text, Me.mskFecFin.Text, gsEmpresa)
        
         
        If MSHFLista.TextMatrix(MSHFLista.Rows - 1, 2) <> "" Then
        MSHFLista.Rows = MSHFLista.Rows + 1
        MSHFLista.TextMatrix(MSHFLista.Rows - 1, 0) = "Total"
        For J = 3 To Me.MSHFLista.Cols - 1
            lnAcumulador = 0
            If Left(MSHFLista.TextMatrix(0, J), 2) <> "U_" And Left(MSHFLista.TextMatrix(0, J), 1) <> "_" Then
                For i = 1 To Me.MSHFLista.Rows - 2
                    If MSHFLista.TextMatrix(i, J) <> "" Then
                        lnAcumulador = lnAcumulador + CCur(MSHFLista.TextMatrix(i, J))
                        
                        If MSHFLista.TextMatrix(i, 8) < 0 Then
                            MSHFLista.Row = i
                            MSHFLista.Col = J
                            MSHFLista.CellBackColor = RGB(300, 150, 150)
                        End If
                        
                    End If
                Next i
                MSHFLista.TextMatrix(MSHFLista.Rows - 1, J) = Format(lnAcumulador, "#,##.00")
                MSHFLista.Row = MSHFLista.Rows - 1
                MSHFLista.Col = J
                MSHFLista.CellBackColor = &HA0C000
                'FlexPrePla.CellFontBold = True
                lnAcumulador = lnAcumulador + CCur(MSHFLista.TextMatrix(i, J))
                
            End If
        Next J
        End If
        MSHFLista.TextMatrix(MSHFLista.Rows - 1, 1) = MSHFLista.Rows - 2
        
        
Case 4

End Select




End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()

Set oReporte = New DRHReportes
Set rs = New ADODB.Recordset

Set oReporte = New DRHReportes
Set rs = oReporte.GetRHReportes
CargaCombo rs, cmbReportes
cmbReportes.ListIndex = 0
Set Progress = New clsProgressBar
mskFecIni.Text = Format(gdFecSis, gsFormatoFechaView)
mskFecFin.Text = Format(gdFecSis, gsFormatoFechaView)
Me.Width = 11985
Me.Height = 8250

End Sub

Private Sub GeneraReporteRH(pflex As MSHFlexGrid, pxlHoja1 As Excel.Worksheet, Optional pnColFiltroVacia As Integer = 0)
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

Sub ReporteQuinta(pgdFecSis As Date, psFecIni As String, psFecFin As String, psEmpresa As String)
 Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim lsCadena As String
    Dim lnMargen As Integer
    Dim lnPagina As Integer
    Dim lnItem As Long
    Dim lsCadAux1 As String
    Dim lsCadAux4 As String
    
    Dim lsProyeccion As String
    Dim lsIngAcumulado As String
    Dim lsImpProyeccion As String
    Dim lsImpAcumulado As String
    Dim lsImpuesto As String
    
    Dim lsCodigo As String * 10
    Dim lsNombre As String * 35
    Dim lsVProy As String * 18
    Dim lsVIngMes As String * 18
    Dim lsVIngAcum As String * 18
    Dim lsVValUIT As String * 18
    Dim lsVIngAfecto As String * 18
    Dim lsVImpuesto As String * 18
    Dim lsVRetencion As String * 18
    Dim lsVImpuestoMes As String * 18
    
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    Dim oInterprete As DInterprete
    Set oInterprete = New DInterprete
    
    Dim lsCadUIT7 As String
    Dim lsCadUIT27 As String
    Dim lsCadUIT54 As String
    Dim lsCadPorHasta27 As String
    Dim lsCadPorHasta54 As String
    Dim lsCadPorMas54 As String
    
    Dim lnCorr As Long
    
    Set rsE = oRep.Rep5taRRHH
    
    'Item
    'Código
    'Apellidos y Nombres
    'Ingreso Mes
    'Ing.Acumulado
    'Ing.Anu.Proy
    'Val.UIT
    'Ing.Afecto
    'Impuesto
    'Impu.Rete
    'Impu.Mes
    
    MSHFLista.Cols = 11
    MSHFLista.TextMatrix(0, 0) = "Item"
    MSHFLista.TextMatrix(0, 1) = "Código"
    MSHFLista.TextMatrix(0, 2) = "Apellidos y Nombres"
    MSHFLista.TextMatrix(0, 3) = "Ing.Mes"
    MSHFLista.TextMatrix(0, 4) = "Ing.Acumul"
    MSHFLista.TextMatrix(0, 5) = "Ing.Anu.Proy"
    MSHFLista.TextMatrix(0, 6) = "Val.UIT"
    MSHFLista.TextMatrix(0, 7) = "Ing.Afecto"
    MSHFLista.TextMatrix(0, 8) = "Impuesto"
    MSHFLista.TextMatrix(0, 9) = "Impu.Rete"
    MSHFLista.TextMatrix(0, 10) = "Impu.Mes"
    
    
    
    lsCadena = ""
    If Not (rsE.EOF And rsE.BOF) Then
        lsCadena = lsCadena & Space(lnMargen) & CentrarCadena(psEmpresa, 180) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & Space(lnMargen) & CentrarCadena("DETALLE DE RETENCIONES-" & Format(pgdFecSis, gsFormatoFechaView), 180) & oImpresora.gPrnSaltoLinea
        oInterprete.Interprete_InI
        lsCadena = lsCadena & Space(lnMargen) & Encabezado("Item;4; ;2;Código;7; ;3;Apellidos y Nombres;23; ;12;Ingreso mes;16; ;2;Ing. Acumulado;17; ;1;Ing.Anu.Proy;15; ;3;Val.UIT;15; ;3;Ing.Afecto;15; ;3;Impuesto;12; ;6;Impu.Rete;15; ;3;Impu.Mes;15; ;3;", lnItem)
        
        lsCadUIT7 = ExprANum(oInterprete.FunEvalua("V_UIT_7", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", ""))
        lsCadUIT27 = ExprANum(oInterprete.FunEvalua("V_UIT_27", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", ""))
        lsCadUIT54 = ExprANum(oInterprete.FunEvalua("V_UIT_54", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", ""))
        lsCadPorHasta27 = oInterprete.FunEvalua("V_POR_IMP_5TA", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", "")
        lsCadPorHasta54 = oInterprete.FunEvalua("V_POR_5TA_H54", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", "")
        lsCadPorMas54 = oInterprete.FunEvalua("V_POR_5TA_M54", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", "")
            
        RSet lsVValUIT = Format(lsCadUIT7, "#.##0.00")
            
            
         oRepEvento_ShowProgress
         lnCorr = 0
        
        While Not rsE.EOF
            lnCorr = lnCorr + 1
            MSHFLista.Rows = MSHFLista.Rows + 1
            lsCodigo = rsE!cRhCod
            lsNombre = PstaNombre(rsE!cPersNombre, False)
            oInterprete.Reinicia
            
            lsCadAux1 = ExprANum(oInterprete.FunEvalua("I_REM_NO_AFEC", rsE!cPersCod, CDate(psFecIni), CDate(psFecFin), False, "", ""))
            lsCadAux4 = Format(oInterprete.GetImp5taEmpRRHH(rsE!cPersCod, CCur(lsCadAux1), CDate(psFecIni), CDate(psFecFin), lsCadUIT7, lsCadUIT27, lsCadUIT54, lsCadPorHasta27, lsCadPorHasta54, lsCadPorMas54, lsCadAux1, lsIngAcumulado, lsProyeccion, lsImpProyeccion, lsImpAcumulado, "", ""), "#0.0000")
            
            RSet lsVIngMes = Format(lsCadAux1, "#.##0.00")
            RSet lsVProy = Format(lsProyeccion, "#.##0.00")
            RSet lsVIngAcum = Format(lsIngAcumulado, "#.##0.00")
            
            If CCur(lsProyeccion) - CCur(lsCadUIT7) < 0 Then
                RSet lsVIngAfecto = Format(0, "#.##0.00")
            Else
                RSet lsVIngAfecto = Format(CCur(lsProyeccion) - CCur(lsCadUIT7), "#.##0.00")
            End If
            lsVImpuesto = FillNum(Format(lsImpProyeccion, "#.##0.00"), 18, " ")
            lsVRetencion = FillNum(Format(lsImpAcumulado, "#.##0.00"), 18, " ")
            lsVImpuestoMes = FillNum(Format(lsCadAux4, "#.##0.00"), 18, " ")
            
            lsCadena = lsCadena & Space(lnMargen) & Format(lnCorr, "0000") & "  " & lsCodigo & lsNombre & lsVIngMes & lsVIngAcum & lsVProy & lsVValUIT & lsVIngAfecto & lsVImpuesto & lsVRetencion & lsVImpuestoMes & oImpresora.gPrnSaltoLinea
            
            MSHFLista.TextMatrix(lnCorr, 0) = Format(lnCorr, "0000")
            MSHFLista.TextMatrix(lnCorr, 1) = lsCodigo
            MSHFLista.TextMatrix(lnCorr, 2) = lsNombre
            MSHFLista.TextMatrix(lnCorr, 3) = Val(lsVIngMes)
            MSHFLista.TextMatrix(lnCorr, 4) = Val(lsVIngAcum)
            MSHFLista.TextMatrix(lnCorr, 5) = Val(lsVProy)
            MSHFLista.TextMatrix(lnCorr, 6) = Val(lsVValUIT)
            MSHFLista.TextMatrix(lnCorr, 7) = Val(lsVIngAfecto)
            MSHFLista.TextMatrix(lnCorr, 8) = Val(lsVImpuesto)
            MSHFLista.TextMatrix(lnCorr, 9) = Val(lsVRetencion)
            MSHFLista.TextMatrix(lnCorr, 10) = Val(lsVImpuestoMes)
            
             oRepEvento_Progress rsE.Bookmark, rsE.RecordCount
            rsE.MoveNext
        Wend
        
        oRepEvento_CloseProgress
    End If
    MSHFLista.Rows = MSHFLista.Rows - 1
    rsE.Close
    Set rsE = Nothing
    Set oRep = Nothing
    Set oInterprete = Nothing
    
    'Rep5taRRHH = lsCadena
End Sub

Private Sub oRepEvento_CloseProgress()
    Progress.CloseForm Me
End Sub

Private Sub oRepEvento_Progress(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Generando Reporte"
End Sub

Private Sub oRepEvento_ShowProgress()
    Progress.ShowForm Me
End Sub



Private Sub mskFecIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    mskFecFin.SetFocus
End If
End Sub
