VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRentaCuart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Renta de Cuarta Categoria"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13845
   Icon            =   "frmRentaCuart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   13845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPDT 
      Caption         =   "PDT  -- >"
      Height          =   405
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   1200
   End
   Begin VB.CommandButton cmdExportaPDT 
      Caption         =   "&Exportar-->"
      Height          =   405
      Left            =   2280
      TabIndex        =   10
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   12480
      TabIndex        =   8
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Excel"
      Height          =   405
      Left            =   11160
      TabIndex        =   7
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "&Procesar"
      Height          =   405
      Left            =   9000
      TabIndex        =   6
      Top             =   480
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
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
      Height          =   765
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   4875
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmRentaCuart.frx":030A
         Left            =   2910
         List            =   "frmRentaCuart.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   1620
      End
      Begin MSMask.MaskEdBox txtAnio 
         Height          =   315
         Left            =   900
         TabIndex        =   2
         Top             =   330
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3765
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   13575
      Begin Sicmact.FlexEdit fe_RepCuarta 
         Height          =   3375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   5953
         Cols0           =   13
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Tipo-Documento-Fecha Emision-Fecha de Pago-Doc-Serie-Nro Doc.-Proveedor-Detalle-Monto-Renta-Pagado"
         EncabezadosAnchos=   "300-1000-1200-1200-1200-1200-1200-1200-2500-2500-1200-1200-1200"
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-C-C-C-C-C-C-L-L-R-R-R"
         FormatosEdit    =   "0-1-1-1-1-1-1-1-1-1-2-2-2"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.OLE OleExcel 
      Class           =   "Excel.Sheet.8"
      Height          =   255
      Left            =   1380
      OleObjectBlob   =   "frmRentaCuart.frx":088A
      TabIndex        =   9
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmRentaCuart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConect As DConecta
Dim dCuarta As New DRegVenta
Dim rs As New ADODB.Recordset
Dim record1 As New ADODB.Recordset
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet


Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdProcesar.SetFocus
    End If
End Sub

Private Sub cmdExportaPDT_Click()

Dim sql As String
On Error GoTo errores
    If ValidaAnio(txtAnio) Then
         Set oConect = New DConecta
         If oConect.AbreConexion() Then
         'INNSERTA DATOS EN LA TABLA RTA4TA
            sql = "Cnt_InsRta4taCatExportar_sp '" & txtAnio.Text & "','" & Format(Right(cboMes.Text, 1), "00") & "'"
            Set record1 = oConect.CargaRecordSet(sql)
         'LEE DATOS DE LA TABLA RTA4TA
            sql = "Cnt_SelRta4taCatExportar_sp"
           Set record1 = oConect.CargaRecordSet(sql)
            oConect.CierraConexion
        End If
        Set oConect = Nothing
        

        If record1.BOF Then
            Set record1 = Nothing
            MsgBox "No existen datos para generar el reporte", vbExclamation, "Aviso!!!"
            Exit Sub
        Else
           record1.MoveFirst
            Call Exporta_FormatoPDT(txtAnio.Text, Format(Right(cboMes.Text, 1), "00"))
           
        End If
    End If
Exit Sub
errores:
 MsgBox Err.Description
End Sub
Private Sub Exporta_FormatoPDT(annio As String, MES As String)
'EXPORTA A WORD
'Dim Row As Integer
'Dim filas_Count As Integer
'Dim mValue As String
'Dim wordApp As Object
'Dim wordDoc As Word.Document
'Dim txtExcelTargetSpec As String
'
'txtExcelTargetSpec = App.path & "\Spooler\0621" & gsNomCmacRUC & annio & MES & ".DOC"
'Set wordApp = New Word.Application
'Set wordDoc = wordApp.Documents.Add
'  wordDoc.Activate
'
'
'     filas_Count = record1.RecordCount
'
'    ' Do While Not record1.EOF
'     For Row = 0 To filas_Count - 1
'        mValue = record1.Fields(0).value
'        wordDoc.ActiveWindow.Selection.InsertAfter mValue
'        wordDoc.ActiveWindow.Selection.InsertParagraphAfter
'        record1.MoveNext
'     Next Row
'    ' Loop
'     Set record1 = Nothing
'     wordDoc.ActiveWindow.Selection.Font.Size = 10
'     wordDoc.ActiveWindow.Selection.EndOf
'
'    wordDoc.SaveAs txtExcelTargetSpec
'    wordDoc.Close
'    wordApp.Quit
'    frmRentaCuart.SetFocus
    'Set wordApp = Nothing
    
    ' exporta a *.txt
    Dim row As Integer
    Dim filas_Count As Integer
    Dim mValue As String
    Dim nfile As Integer
    filas_Count = record1.RecordCount
    nfile = FreeFile
    Open App.path & "\Spooler\0621" & gsNomCmacRUC & annio & MES & ".TXT" For Output As #nfile
        For row = 0 To filas_Count - 1
            mValue = record1.Fields(0).value
            Print #nfile, mValue
            record1.MoveNext
         Next row
    Close #nfile
     MsgBox ("Archivo: 0621" & gsNomCmacRUC & annio & MES & ".TXT" & " Generado Correctamente")

End Sub
Private Sub cmdExportar_Click()
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    'If Me.MSHLista.TextMatrix(1, 1) = "" Then
    If rs.RecordCount = 0 Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    'ORCR 20140217****************************************
    'lsArchivoN = App.path & "\Spooler\" & "RenCuarta" & Format(gdFecSis, "yyyy") & Format(Right(cboMes.Text, 1), "00") & ".xls"
    lsArchivoN = App.path & "\Spooler\" & "RenCuarta" & Format(gdFecSis, "yyyy") & Format(Right(cboMes.Text, 1), "00") & ".xlsx"
    '*****************************************************
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       GeneraReporteCuarta xlHoja1 'MSHLista quitado por pasi
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

Private Sub cmdPDT_Click()
    frmArchRecibosHonorarios.Show 1
End Sub

Private Sub cmdProcesar_Click()
    Dim row As Integer 'Agregado PASI20140205
   If ValidaAnio(txtAnio) Then
       Set rs = dCuarta.CargaCuarta(txtAnio.Text, Format(Right(cboMes.Text, 2), "00"))
      'Modificado PASI20140205
      'Set MSHLista.DataSource = rs
      FormateaFlex fe_RepCuarta
      If Not rs.EOF Then
        Do While Not rs.EOF
            fe_RepCuarta.AdicionaFila
            row = fe_RepCuarta.row
            fe_RepCuarta.TextMatrix(row, 1) = rs!Tipo
            fe_RepCuarta.TextMatrix(row, 2) = rs!Documento
            fe_RepCuarta.TextMatrix(row, 3) = rs!FECHAEMISION
            fe_RepCuarta.TextMatrix(row, 4) = rs!FechaPago
            fe_RepCuarta.TextMatrix(row, 5) = rs!DOC
            fe_RepCuarta.TextMatrix(row, 6) = rs!Serie
            fe_RepCuarta.TextMatrix(row, 7) = rs!NRODOCUMENTO
            fe_RepCuarta.TextMatrix(row, 8) = rs!Proveedor
            fe_RepCuarta.TextMatrix(row, 9) = rs!Detalle
            fe_RepCuarta.TextMatrix(row, 10) = rs!Monto
            fe_RepCuarta.TextMatrix(row, 11) = rs!Renta
            fe_RepCuarta.TextMatrix(row, 12) = rs!PAGADO
            rs.MoveNext
        Loop
      Else
        MsgBox "No hay Datos para Mostrar", vbInformation, "Aviso!!!"
      End If
      'end PASI
      cmdExportaPDT.Enabled = True
    cmdExportar.Enabled = True
   End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
   CentraForm Me
   txtAnio = Year(gdFecSis)
   cboMes.ListIndex = Month(gdFecSis) - 1
    Set dCuarta = New DRegVenta
    Set rs = New ADODB.Recordset
    ConfigurarMSHLista
    cmdExportaPDT.Enabled = False
    cmdExportar.Enabled = False
End Sub


Private Sub txtAnio_KeyPress(KeyAscii As Integer)
   KeyAscii = NumerosEnteros(KeyAscii)
   If KeyAscii = 13 Then
       cboMes.SetFocus
   End If
End Sub

'Public Sub GeneraReporteCuarta(pflex As MSHFlexGrid, pxlHoja1 As Excel.Worksheet, Optional pnColFiltroVacia As Integer = 0) 'Comentado PASI20140506
Public Sub GeneraReporteCuarta(pxlHoja1 As Excel.Worksheet, Optional pnColFiltroVacia As Integer = 0)
    Dim I As Integer
    Dim k As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    Dim nTotal As Currency
    Dim nRenta As Currency
    Dim nPagado As Currency
    Dim Lineas As Integer
        
    'Modificado PASI20140205
    xlAplicacion.Range("A1:A1").ColumnWidth = 12
    xlAplicacion.Range("B1:B1").ColumnWidth = 15
    xlAplicacion.Range("c1:c1").ColumnWidth = 15
    xlAplicacion.Range("D1:D1").ColumnWidth = 15
    xlAplicacion.Range("E1:E1").ColumnWidth = 15
    xlAplicacion.Range("F1:F1").ColumnWidth = 40
    xlAplicacion.Range("G1:G1").ColumnWidth = 15
    xlAplicacion.Range("H1:H1").ColumnWidth = 15
    xlAplicacion.Range("I1:I1").ColumnWidth = 60
    xlAplicacion.Range("J1:J1").ColumnWidth = 80
    xlAplicacion.Range("K1:K1").ColumnWidth = 30
    xlAplicacion.Range("L1:L1").ColumnWidth = 30
    xlAplicacion.Range("L1:L1").ColumnWidth = 30
    xlAplicacion.Range("M1:M1").ColumnWidth = 30
    'end PASI

    Dim lFechaMes As String
    lFechaMes = Choose(Month(gdFecSis), "Enero", "Febrero", "Marzo", "Abril", _
                                        "Mayo", "Junio", "Julio", "Agosto", _
                                        "Setiembre", "Octubre", "Noviembre", "Diciembre")
    
    pxlHoja1.Cells(2, 3) = " LIBRO  DE  RETENCIONES  MES  DE" & "   " & UCase(lFechaMes) & "  DEL  " & Year(gdFecSis)
    xlHoja1.Range(xlHoja1.Cells(2, 3), xlHoja1.Cells(2, 7)).Merge True
    pxlHoja1.Range(xlHoja1.Cells(2, 3), xlHoja1.Cells(2, 7)).Font.Bold = True
         
    Lineas = 4
    
    'Modificaso PASI20140205
'    For i = 0 To pflex.Rows - 1
'        If pnColFiltroVacia = 0 Then
'            For j = 0 To pflex.Cols - 1
''                If I >= 1 And j = 1 Then
''                   pxlHoja1.Cells(Lineas + 1, j + 1) = Format(pflex.TextMatrix(I, j), "mm/dd/yy")
''                Else
''                   pxlHoja1.Cells(Lineas + 1, j + 1) = pflex.TextMatrix(I, j)
''                End If
''                If j = 5 And I > 0 Then
''                    nTotal = nTotal + Val(pflex.TextMatrix(I, 6))
''                End If
''                If j = 6 And I > 0 Then
''                    nRenta = nRenta + Val(pflex.TextMatrix(I, 7))
''                End If
''                If j = 7 And I > 0 Then
''                    nPagado = nPagado + Val(pflex.TextMatrix(I, 8))
''                End If
'
'                    'Modificado PASI20140205
'                    'pxlHoja1.Cells(Lineas + 1, j + 1) = pflex.TextMatrix(i, j) 'EJVG20121110
'                    If j = 3 Then
'                    pxlHoja1.Cells(Lineas + 1, j + 1) = CStr(pflex.TextMatrix(i, j))
'                    Else
'                    pxlHoja1.Cells(Lineas + 1, j + 1) = pflex.TextMatrix(i, j) 'EJVG20121110
'                   End If
'
'            Next j
'            Lineas = Lineas + 1
'        Else
'            If pflex.TextMatrix(i, pnColFiltroVacia) <> "" Then
'                For j = 0 To pflex.Cols - 1
'                    pxlHoja1.Cells(Lineas + 1, j + 1) = pflex.TextMatrix(i, j)
'                Next j
'            End If
'        End If
'    Next i
    
    pxlHoja1.Range("D:E").NumberFormat = "dd/mm/yyyy" 'Agregado PASI20140502
    
    For I = 0 To fe_RepCuarta.Rows - 1
        If pnColFiltroVacia = 0 Then
            For j = 0 To fe_RepCuarta.Cols - 1
'                If I >= 1 And j = 1 Then
'                   pxlHoja1.Cells(Lineas + 1, j + 1) = Format(pflex.TextMatrix(I, j), "mm/dd/yy")
'                Else
'                   pxlHoja1.Cells(Lineas + 1, j + 1) = pflex.TextMatrix(I, j)
'                End If
'                If j = 5 And I > 0 Then
'                    nTotal = nTotal + Val(pflex.TextMatrix(I, 6))
'                End If
'                If j = 6 And I > 0 Then
'                    nRenta = nRenta + Val(pflex.TextMatrix(I, 7))
'                End If
'                If j = 7 And I > 0 Then
'                    nPagado = nPagado + Val(pflex.TextMatrix(I, 8))
'                End If

                    'Modificado PASI20140205
                    'pxlHoja1.Cells(Lineas + 1, j + 1) = pflex.TextMatrix(i, j) 'EJVG20121110
                    If (j = 3 Or j = 4) And I > 0 Then
                    'pxlHoja1.Range(xlHoja1.Cells(2, 3), xlHoja1.Cells(2, 7))
                    
                    pxlHoja1.Cells(Lineas + 1, j + 1) = CDate(fe_RepCuarta.TextMatrix(I, j))
                    Else
                    pxlHoja1.Cells(Lineas + 1, j + 1) = fe_RepCuarta.TextMatrix(I, j) 'EJVG20121110
                   End If
                
            Next j
            Lineas = Lineas + 1
        Else
            If fe_RepCuarta.TextMatrix(I, pnColFiltroVacia) <> "" Then
                For j = 0 To fe_RepCuarta.Cols - 1
                    pxlHoja1.Cells(Lineas + 1, j + 1) = fe_RepCuarta.TextMatrix(I, j)
                Next j
            End If
        End If
    Next I
    'End PASI
    
    'pxlHoja1.Cells(Lineas + 1, 7) = nTotal
    'pxlHoja1.Cells(Lineas + 1, 8) = nRenta
    'pxlHoja1.Cells(Lineas + 1, 9) = nPagado
    
    'Modificado PASI20140205
'    pxlHoja1.Cells(Lineas + 1, 9) = "=SUM(J6:J" & Lineas & ")"
'    pxlHoja1.Cells(Lineas + 1, 10) = "=SUM(K6:K" & Lineas & ")"
'    pxlHoja1.Cells(Lineas + 1, 11) = "=SUM(L6:L" & Lineas & ")"
    
    pxlHoja1.Cells(Lineas + 1, 11) = "=SUM(K6:K" & Lineas & ")"
    pxlHoja1.Cells(Lineas + 1, 12) = "=SUM(L6:L" & Lineas & ")"
    pxlHoja1.Cells(Lineas + 1, 13) = "=SUM(M6:M" & Lineas & ")"
    
    
    pxlHoja1.Range("K6:M" & Lineas + 1).NumberFormat = "#,##0.00"
    pxlHoja1.Range("C:C").HorizontalAlignment = xlCenter
    'end PASI
    
    'Modificado PASI20140205
'    pxlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 11)).Font.Bold = True
'    pxlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 11)).HorizontalAlignment = xlCenter
'    pxlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 11)).VerticalAlignment = xlCenter
'    pxlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 11)).Borders.LineStyle = 1
'    pxlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 11)).Interior.ColorIndex = 36 '.Color = RGB(159, 206, 238)
'
'
'    pxlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 11)).EntireRow.AutoFit
'    pxlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 11)).WrapText = True
    
    pxlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 13)).Font.Bold = True
    pxlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 13)).HorizontalAlignment = xlCenter
    pxlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 13)).VerticalAlignment = xlCenter
    pxlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 13)).Borders.LineStyle = 1
    pxlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 13)).Interior.ColorIndex = 36 '.Color = RGB(159, 206, 238)


    pxlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 13)).EntireRow.AutoFit
    pxlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 13)).WrapText = True
    
    'end PASI
    
End Sub

Sub ConfigurarMSHLista()
    'With MSHLista
'        .Cols = 3
'        .Rows = 2
'        MSH.TextMatrix(0, 0) = "ASIGNACION"
'        MSH.TextMatrix(0, 1) = "NRO"
'        MSH.TextMatrix(0, 2) = "DOCUMENTO"
'        MSHLista.ColWidth(0) = 2300
'        MSHLista.ColWidth(1) = 1800
'        MSHLista.ColWidth(2) = 1600
    'End With
End Sub
