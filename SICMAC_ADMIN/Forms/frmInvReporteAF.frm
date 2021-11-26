VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInvReporteAF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Activos Fijos"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8490
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvReporteAF.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de Búsqueda"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.Frame Frame3 
         Caption         =   "Compras"
         Height          =   615
         Left            =   4080
         TabIndex        =   18
         Top             =   1440
         Width           =   3495
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   480
            TabIndex        =   19
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   17367041
            CurrentDate     =   31048
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   2040
            TabIndex        =   20
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   17367041
            CurrentDate     =   40543
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Al:"
            Height          =   195
            Left            =   1800
            TabIndex        =   22
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Del:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   360
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Mostrar Bienes"
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   3735
         Begin VB.OptionButton optNoDepre 
            Caption         =   "No Depreciables"
            Height          =   255
            Left            =   1800
            TabIndex        =   17
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optSiDepre 
            Caption         =   "Depreciables"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         ItemData        =   "frmInvReporteAF.frx":030A
         Left            =   5400
         List            =   "frmInvReporteAF.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Width           =   2820
      End
      Begin VB.TextBox txtPeriodo 
         Height          =   285
         Left            =   5400
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   6360
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1200
         TabIndex        =   6
         Top             =   390
         Width           =   930
      End
      Begin VB.ComboBox cmbBien 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   3300
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   3300
      End
      Begin Sicmact.TxtBuscar TxtAgencia 
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   360
         Width           =   855
         _ExtentX        =   1296
         _ExtentY        =   503
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin VB.Label lblMes 
         Caption         =   "Mes :"
         Height          =   210
         Left            =   4665
         TabIndex        =   14
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label3 
         Caption         =   "Periodo:"
         Height          =   255
         Left            =   4560
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.OLE OleExcel 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         Enabled         =   0   'False
         Height          =   240
         Left            =   120
         SizeMode        =   1  'Stretch
         TabIndex        =   10
         Top             =   2280
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "Bien:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Agencia:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmInvReporteAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oInventario As NInvActivoFijo
Dim oArea As DActualizaDatosArea
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Private Sub cmbTipo_Click()
    Dim rs As ADODB.Recordset
    Dim ii As Integer
    Set oInventario = New NInvActivoFijo
    Set rs = oInventario.ObtenerBienXTipo(Right(cmbTipo.Text, 5))
    cmbBien.Clear
    For ii = 0 To rs.RecordCount - 1
        cmbBien.AddItem rs.Fields(0) & Space(50) & "," & rs.Fields(1)
        rs.MoveNext
    Next ii
    cmbBien.AddItem "TODOS", 0
    cmbBien.Text = "TODOS"
End Sub

Private Function DevolverCodBien(ByVal sCod As String)
    Dim liPosicion As Integer
    Dim lsCod As String
    lsCod = Trim(sCod)
    liPosicion = InStr(lsCod, ",")
    If liPosicion > 0 Then
    DevolverCodBien = Mid(lsCod, liPosicion + 1, Len(lsCod))
    End If
    DevolverCodBien = DevolverCodBien
End Function

Private Sub Command1_Click()
    Dim rsDatos As ADODB.Recordset
    Set rsDatos = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    Set oInventario = New NInvActivoFijo
    Dim lsCategoBien As String '*** PEAC 20100507
    Dim ldFecha As Date
    
    If cmbMes.Text <> "" Then
        'ldFecha = CDate(Trim(Format(Trim(Right(Me.cmbMes.Text, 5)), "00") & "/" & "01" & "/" & txtPeriodo.Text))
        ldFecha = CDate("01" & "/" & Trim(Format(Trim(Right(Me.cmbMes.Text, 5)), "00") & "/" & txtPeriodo.Text))
    Else
        MsgBox "Debe escoger Mes", vbCritical
        Exit Sub
    End If
    
    '*** PEAC 20110806
    If Me.DTPicker1.value > Me.DTPicker2.value Then
        MsgBox "Verifique las fechas por favor.", vbCritical + vbOKOnly, "Atención"
        Exit Sub
    End If
    
    lsCategoBien = IIf(Me.optSiDepre.value = True, "1", "0") '*** PEAC 20100507 - (1=DEPRECIABLE, 0=NO DEPRECIABLE)
        
    '***PEAC 20100507 - SE AGREGO PARAMETRO (lsCategoBien)
    '*** PEAC 20110806 - SE AGREGO RANGO DE FECHAS
    Set rsDatos = oInventario.ObtenerReporteAF(IIf(chkTodos.value = 1, "", TxtAgencia.Text), IIf(cmbTipo.Text = "TODOS", "", Right(cmbTipo.Text, 5)), DevolverCodBien(cmbBien.Text), ldFecha, lsCategoBien, Format(DTPicker1.value, "yyyymmdd"), Format(DTPicker2.value, "yyyymmdd"))
    ', IIf(cmbBien.Text = "TODOS", "", Right(cmbBien.Text, 5)))
    
    If rsDatos Is Nothing Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lsArchivoN = App.path & "\Spooler\" & "ReporteActivoFijo" & Format(CDate(gdFecSis), "yyyymmdd") & Format(Time, "hhmmss") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       
       'ReporteAFCabeceraExcel xlHoja1 '*** PEAC 20110808
       ReporteAFCabeceraExcel xlHoja1, Me.optSiDepre.value
       
       'GeneraReporte rsDatos '*** PEAC 20110808
       GeneraReporte rsDatos, Me.optSiDepre.value, ldFecha
       
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
    MousePointer = 0
End Sub

Private Function DepreAcumuladaEjeAnt(ByVal sFInicio As String, ByVal sVidaUtil As String, ByVal pdFecFin As Date) As String
    Dim i As Integer
    Dim meses As Integer
    
    'For i = Mid(sFInicio, 7, 4) To Trim(Mid("31/01/2009", 7, 4) - 1) '*** PEAC 20110816
    For i = Mid(sFInicio, 7, 4) To Trim(Str(Year(pdFecFin) - 1))
        If i = Mid(sFInicio, 7, 4) Then
            meses = meses + DateDiff("m", CDate(sFInicio), CDate("31/12/" & Str(i)))
        Else
            meses = meses + DateDiff("m", CDate("01/12/" & Str(i - 1)), CDate("31/12/" & Str(i)))
        End If
    Next i
    DepreAcumuladaEjeAnt = meses
End Function

Private Function DepreEjercicio() As String
    Dim i As Integer
    Dim MesesEjercicio As Integer
        MesesEjercicio = DateDiff("m", CDate("01/12/" & Str(Year(Date) - 1)), CDate("31/01/2009"))
    DepreEjercicio = MesesEjercicio
End Function

Private Sub GeneraReporte(prRs As ADODB.Recordset, Optional pbDepre As Boolean = True, Optional pdFecFin As Date = "01/01/1900")
    Dim i As Integer
    Dim J As Integer
    
    Dim lnSI As Currency '(6) Saldo Inicial
    Dim lnVH As Currency '(11) Valor Historico
    Dim lnVA As Currency '(13) Valor Ajustado
    Dim lnDACEA As Currency '(19) Depre Acum Cier Ejer Ant
    Dim lnDE As Currency '(20) Depre Ejerc
    
    Dim lnDEx As Currency ''(20) Depre Ejerc - PEAC 20110816
    
    Dim lnDAH As Currency '(23) Depre Acum Hist
    Dim lnDAAI As Currency '(25) Depre Acum Ajust Inflac
    
    Dim lnNumMesesEjer As Integer '*** PEAC 20110816
    
    
    '*** PEAC 20120326
    i = 8

    While Not prRs.EOF

        i = i + 1

        xlHoja1.Cells(i, 2) = prRs!cSerie
        xlHoja1.Cells(i, 3) = prRs!CtaCont
        xlHoja1.Cells(i, 4) = prRs!cDescripcion
        xlHoja1.Cells(i, 5) = prRs!vMarca
        xlHoja1.Cells(i, 6) = prRs!vModelo
        xlHoja1.Cells(i, 7) = prRs!vSerie
        xlHoja1.Cells(i, 8) = IIf(Format(prRs!dCompra, "yyyymmdd") <= prRs!cUltDiaAnioAnterior, prRs!nBSValor, 0#)
        xlHoja1.Cells(i, 9) = IIf(Format(prRs!dCompra, "yyyymmdd") > prRs!cUltDiaAnioAnterior, prRs!nBSValor, 0#)
        xlHoja1.Cells(i, 10) = "'" + Format(prRs!dCompra, "dd/mm/yyyy")
        xlHoja1.Cells(i, 11) = "'" + Format(prRs!dActivacion, "dd/mm/yyyy")
        xlHoja1.Cells(i, 12) = prRs!cMetodoApli
        xlHoja1.Cells(i, 14) = prRs!PorcentajeDepreciacion
        xlHoja1.Cells(i, 15) = Format(prRs!nDepreDelEjerAnt, "#,#.00")
        xlHoja1.Cells(i, 16) = Format(prRs!nDepreDelEjer, "#,#.00")
        xlHoja1.Cells(i, 17) = Format(prRs!nDepreDelEjerAnt + prRs!nDepreDelEjer, "#,#.00")

        prRs.MoveNext
    Wend

Exit Sub
'*** FIN PEAC
        
        
    i = 8
    'prRs.MoveFirst
    While Not prRs.EOF
        i = i + 1
        For J = 0 To prRs.Fields.Count - 1

            If Not pbDepre And J = 16 Then  '*** PEAC 20110808
                Exit For
            End If
            
            If IsNumeric(prRs.Fields(J)) And (J = 6 Or J = 11 Or J = 13 Or J = 23 Or J = 25) Then ' Then

                xlHoja1.Cells(i + 1, J + 1) = Format(prRs.Fields(J), "#,##0.00")
                
                Select Case J
                Case 6
                    lnSI = lnSI + CCur(prRs.Fields(J))
                Case 11
                    lnVH = lnVH + CCur(prRs.Fields(J))
                Case 13
                    lnVA = lnVA + CCur(prRs.Fields(J))
                Case 23
                    lnDAH = lnDAH + CCur(prRs.Fields(J))
                Case 25
                    lnDAAI = lnDAAI + CCur(prRs.Fields(J))
                End Select
            Else
                If J = 19 Then
                    If prRs!nBSPerDeprecia <> prRs!PeriodosDeprecia Then
                    
                        xlHoja1.Cells(i + 1, J + 1) = Format(Round((prRs!nBSValor / prRs!nBSPerDeprecia) * DepreAcumuladaEjeAnt(prRs!FAdquisicion, prRs!nBSPerDeprecia, pdFecFin), 2), "#,##0.00")
                        ''xlHoja1.Cells(i + 1, J + 1) = Format(prRs!nDepreDelEjerAnt, "#,##0.00")
                    Else
                        xlHoja1.Cells(i + 1, J + 1) = Format(prRs!nBSValor, "#,##0.00")
                    End If
                    lnDACEA = lnDACEA + CCur(xlHoja1.Cells(i + 1, J + 1))
                Else
                    If J = 20 Then
                        If prRs!nBSPerDeprecia <> prRs!PeriodosDeprecia Then
                            lnNumMesesEjer = DateDiff("m", CDate("31/12/" & Str(Year(pdFecFin) - 1)), pdFecFin) ''CDate("31/01/2009"))
'                            xlHoja1.Cells(i + 1, J + 1) = Format(Round((prRs!nBSValor / prRs!nBSPerDeprecia) * DepreEjercicio, 2), "#,##0.00")
'                            lnDE = lnDE + CCur(Format(Round((prRs!nBSValor / prRs!nBSPerDeprecia) * DepreEjercicio, 2), "#,##0.00"))

                            xlHoja1.Cells(i + 1, J + 1) = Format(Round((prRs!nBSValor / prRs!nBSPerDeprecia) * lnNumMesesEjer, 2), "#,##0.00")
                            lnDE = lnDE + CCur(Format(Round((prRs!nBSValor / prRs!nBSPerDeprecia) * lnNumMesesEjer, 2), "#,##0.00"))
                        
                        End If
                    Else
                                                
                        xlHoja1.Cells(i + 1, J + 1) = prRs.Fields(J)
                    End If
                End If
            End If

        Next J
        
        If (xlHoja1.Cells(i + 1, 24) > 0 And xlHoja1.Cells(i + 1, 20) > 0) And (xlHoja1.Cells(i + 1, 24) <> xlHoja1.Cells(i + 1, 20)) Then
            xlHoja1.Cells(i + 1, 21) = xlHoja1.Cells(i + 1, 24) - xlHoja1.Cells(i + 1, 20)
            
            lnDEx = lnDEx + Format(xlHoja1.Cells(i + 1, 24) - xlHoja1.Cells(i + 1, 20), "#,##0.00")
            
        ElseIf (xlHoja1.Cells(i + 1, 24) > 0 And xlHoja1.Cells(i + 1, 20) = 0) Then
            xlHoja1.Cells(i + 1, 21) = xlHoja1.Cells(i + 1, 24)
            
            lnDEx = lnDEx + Format(xlHoja1.Cells(i + 1, 24), "#,##0.00")
            
        End If
        
        prRs.MoveNext
    Wend
    
    xlHoja1.Cells(prRs.RecordCount + 10, 6) = "TOTALES"
    xlHoja1.Cells(prRs.RecordCount + 10, 7) = Format(lnSI, "#,##0.00")
    xlHoja1.Cells(prRs.RecordCount + 10, 12) = Format(lnVH, "#,##0.00")
    xlHoja1.Cells(prRs.RecordCount + 10, 14) = Format(lnVA, "#,##0.00")
    
    If pbDepre Then '*** PEAC 20110808
    
        xlHoja1.Cells(prRs.RecordCount + 10, 20) = Format(lnDACEA, "#,##0.00")
        'xlHoja1.Cells(prRs.RecordCount + 10, 21) = Format(lnDE, "#,##0.00")
        xlHoja1.Cells(prRs.RecordCount + 10, 21) = Format(lnDEx, "#,##0.00")
        xlHoja1.Cells(prRs.RecordCount + 10, 24) = Format(lnDAH, "#,##0.00")
        xlHoja1.Cells(prRs.RecordCount + 10, 26) = Format(lnDAAI, "#,##0.00")
    
    End If
    
    xlHoja1.Range("B10:B" & prRs.RecordCount + 9).HorizontalAlignment = xlCenter
    xlHoja1.Range("O10:O" & prRs.RecordCount + 9).HorizontalAlignment = xlCenter
    xlHoja1.Range("P10:P" & prRs.RecordCount + 9).HorizontalAlignment = xlCenter
    xlHoja1.Range("Q10:Q" & prRs.RecordCount + 9).HorizontalAlignment = xlCenter
    xlHoja1.Range("S10:S" & prRs.RecordCount + 9).HorizontalAlignment = xlCenter
    xlHoja1.Range("V10:V" & prRs.RecordCount + 9).HorizontalAlignment = xlCenter
    
    If pbDepre Then '*** PEAC 20110808
        'Border's Tabla
        xlHoja1.Range("A10:Z" & prRs.RecordCount + 9).BorderAround xlContinuous, xlMedium
        xlHoja1.Range("A10:Z" & prRs.RecordCount + 9).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        xlHoja1.Range("A10:Z" & prRs.RecordCount + 9).Borders(xlInsideVertical).LineStyle = xlContinuous
    Else
        xlHoja1.Range("A10:P" & prRs.RecordCount + 9).BorderAround xlContinuous, xlMedium
        xlHoja1.Range("A10:P" & prRs.RecordCount + 9).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        xlHoja1.Range("A10:P" & prRs.RecordCount + 9).Borders(xlInsideVertical).LineStyle = xlContinuous
    End If
    
    If pbDepre Then '*** PEAC 20110808
        'Border's Totales
        xlHoja1.Range("G" & prRs.RecordCount + 10 & ":Z" & prRs.RecordCount + 10).BorderAround xlContinuous, xlMedium
        xlHoja1.Range("G" & prRs.RecordCount + 10 & ":Z" & prRs.RecordCount + 10).Borders(xlInsideVertical).LineStyle = xlContinuous
    End If
    
    xlHoja1.Range("J8:J9").Cells.VerticalAlignment = xlJustify
    
    xlHoja1.Range("L8:L9").Cells.VerticalAlignment = xlJustify
    
    xlHoja1.Range("H8:H9").Cells.VerticalAlignment = xlJustify
    xlHoja1.Range("N8:N9").Cells.VerticalAlignment = xlJustify
    xlHoja1.Range("M8:M9").Cells.VerticalAlignment = xlJustify
    xlHoja1.Range("O8:O9").Cells.VerticalAlignment = xlJustify
    xlHoja1.Range("P8:P9").Cells.VerticalAlignment = xlJustify
    xlHoja1.Range("Q9:Q9").Cells.VerticalAlignment = xlJustify
    xlHoja1.Range("R9:R9").Cells.VerticalAlignment = xlJustify
    xlHoja1.Range("V8:V9").Cells.VerticalAlignment = xlJustify
    xlHoja1.Range("S8:S9").Cells.VerticalAlignment = xlJustify
    xlHoja1.Range("T8:T9").Cells.VerticalAlignment = xlJustify
    xlHoja1.Range("U8:U9").Cells.VerticalAlignment = xlJustify
    xlHoja1.Range("W8:W9").Cells.VerticalAlignment = xlJustify
    xlHoja1.Range("X8:X9").Cells.VerticalAlignment = xlJustify
    xlHoja1.Range("Y8:Y9").Cells.VerticalAlignment = xlJustify
    xlHoja1.Range("Z8:Z9").Cells.VerticalAlignment = xlJustify
End Sub

Public Function ReporteAFCabeceraExcel(Optional xlHoja1 As Excel.Worksheet, Optional pbDepre As Boolean = True) As String
    xlHoja1.PageSetup.LeftMargin = 1.5
    xlHoja1.PageSetup.RightMargin = 0
    xlHoja1.PageSetup.BottomMargin = 1
    xlHoja1.PageSetup.TopMargin = 1
    xlHoja1.PageSetup.Zoom = 70
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 8


    '*** PEAC 20120326
    xlHoja1.Cells(2, 2) = "FORMATO 7.1 : ''REGISTRO DE ACTIVOS FIJOS - DETALLE DE LOS ACTIVOS FIJOS''"
    xlHoja1.Cells(4, 2) = "PERIODO : " + txtPeriodo.Text
    xlHoja1.Cells(5, 2) = "RUC : 20103845328"
    xlHoja1.Cells(6, 2) = "APELLIDOS Y NOMBRES, DENOMINACION O RAZÓN SOCIAL : CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS SA."

    xlHoja1.Range("B7:B8").MergeCells = True
    xlHoja1.Range("C7:C8").MergeCells = True
    xlHoja1.Range("D7:G7").MergeCells = True
    xlHoja1.Range("H7:H8").MergeCells = True
    xlHoja1.Range("I7:I8").MergeCells = True
    xlHoja1.Range("J7:J8").MergeCells = True
    xlHoja1.Range("K7:K8").MergeCells = True
    xlHoja1.Range("L7:M7").MergeCells = True
    xlHoja1.Range("N7:N8").MergeCells = True
    xlHoja1.Range("O7:O8").MergeCells = True
    xlHoja1.Range("P7:P8").MergeCells = True
    xlHoja1.Range("Q7:Q8").MergeCells = True

    xlHoja1.Cells(7, 2) = "CODIGO RELACIONADO CON EL ACTIVO FIJO"
    xlHoja1.Cells(7, 3) = "CUENTA CONTA-BLE DEL ACTIVO FIJO"
    xlHoja1.Cells(7, 4) = "DETALLE DE L ACTIVO FIJO"

    xlHoja1.Cells(8, 4) = "DESCRIPCIÓN: MAQUINARIAS"
    xlHoja1.Cells(8, 5) = "MARCA DEL ACTIVO FIJO"
    xlHoja1.Cells(8, 6) = "MODELO DEL ACTIVO FIJO"
    xlHoja1.Cells(8, 7) = "NUMERO DE SERIE Y/O PLACA DEL ACTIVO FIJO"
    xlHoja1.Cells(7, 8) = "SALDO INICIAL"
    xlHoja1.Cells(7, 9) = "ADQUISI-CIONES ADICIONES"
    xlHoja1.Cells(7, 10) = "FECHA DE ADQUISI-CIÓN"
    xlHoja1.Cells(7, 11) = "FECHA DE INICIO DEL USO DEL ACTIVO FIJO"
    xlHoja1.Cells(7, 12) = "DEPRECIACION"
    xlHoja1.Cells(8, 12) = "METODO APLICADO"
    xlHoja1.Cells(8, 13) = "N° DE DOCUMEN-TO DE AUTORIZA-CIÓN"
    xlHoja1.Cells(7, 14) = "PORCEN-TAJE DE DEPRE-CIACIÓN"
    xlHoja1.Cells(7, 15) = "DEPRECIA-CIÓN ACUMULA-DA AL CIERRE DEL EJERCICIO ANTERIOR"
    xlHoja1.Cells(7, 16) = "DEPRECIA-CIÓN DEL EJERCICIO"
    xlHoja1.Cells(7, 17) = "DEPRECIA-CIÓN ACUMULADA HÍSTORICA"

    xlHoja1.Rows("8:8").RowHeight = 77.25
    xlHoja1.Columns("B:B").ColumnWidth = 16
    xlHoja1.Columns("C:C").ColumnWidth = 14
    xlHoja1.Columns("D:D").ColumnWidth = 38

    xlHoja1.Range("B7:Q8").HorizontalAlignment = xlCenter
    xlHoja1.Range("B7:Q8").VerticalAlignment = xlCenter
    xlHoja1.Range("B2:Q8").Font.Bold = True
    xlHoja1.Range("B7:Q8").WrapText = True
    xlHoja1.Range("B7:Q8").Borders.LineStyle = 1 ''(xlDiagonalDown).LineStyle = xlNone


'    ApExcel.Cells.Select
'    ApExcel.Cells.EntireColumn.AutoFit
'    ApExcel.Columns("B:B").ColumnWidth = 6#
'    ApExcel.Range("B2").Select



Exit Function
    '*** FIN PEAC
    
    
    xlHoja1.Cells(3, 1) = "AGENCIA: "
    xlHoja1.Cells(3, 2) = IIf(chkTodos.value = 1, "TODOS", lblAgencia.Caption)
    xlHoja1.Cells(3, 4) = "TIPO: "
    xlHoja1.Cells(3, 5) = Mid(cmbTipo.Text, 1, 30)
    xlHoja1.Cells(3, 7) = "BIEN: "
    xlHoja1.Cells(3, 8) = Mid(cmbBien.Text, 1, 30)
    
    xlHoja1.Cells(2, 12) = "REPORTE DE ACTIVOS FIJOS" & " " & Mid(cmbMes.Text, 1, 8) & " " & txtPeriodo.Text
    
    xlHoja1.Cells(4, 1) = "PERIODO: "
    xlHoja1.Cells(4, 2) = txtPeriodo.Text
    
    xlHoja1.Cells(5, 1) = "R.U.C."
    xlHoja1.Cells(5, 2) = "20103845328"
    
    xlHoja1.Cells(6, 1) = "DENOMINACIÓN: "
    xlHoja1.Cells(6, 2) = "CMAC MAYNAS S.A."
    
    xlHoja1.Cells(8, 1) = "CÓDIGO: "
    xlHoja1.Cells(8, 2) = "CTA CONTABLE"
    
    xlHoja1.Cells(8, 3) = "DETALLE DEL ACTIVO FIJO"
    
    xlHoja1.Cells(9, 3) = "DESCRIPCIÓN"
    xlHoja1.Cells(9, 4) = "MARCA"
    xlHoja1.Cells(9, 5) = "MODELO"
    xlHoja1.Cells(9, 6) = "SERIE"
    
    xlHoja1.Cells(8, 7) = "SALDO INICIAL"
    
    xlHoja1.Cells(8, 8) = "ADQUISICIONES ADICIONES"
    xlHoja1.Cells(8, 9) = "MEJORAS"
    xlHoja1.Cells(8, 10) = "RETIROS Y/O BAJAS"
    xlHoja1.Cells(8, 11) = "OTROS AJUSTES"
    xlHoja1.Cells(8, 12) = "VALOR HISTORICO DEL ACTIVO FIJO AL: " & Year(Date) - 1
    
    xlHoja1.Cells(8, 13) = "AJUSTE POR INFLACION"
    xlHoja1.Cells(8, 14) = "VALOR AJUSTADO DEL ACTIVO FIJO AL: " & Year(Date) - 1
    xlHoja1.Cells(8, 15) = "FECHA DE ADQUISICION"
    xlHoja1.Cells(8, 16) = "FECHA DE INICIO DEL USO DEL ACTIVO FIJO"
    
    If pbDepre Then '*** PEAC 20110808
    
        xlHoja1.Cells(9, 17) = "METODO APLICADO"
        xlHoja1.Cells(9, 18) = "N° DE DOCUMENTO DE AUTORIZACION"
        xlHoja1.Cells(8, 17) = "DEPRECIACION"
        xlHoja1.Cells(9, 19) = "PORCENTAJE DE LA DEPRECIACION"
        xlHoja1.Cells(9, 20) = "DEPRECIACION ACUMULADA AL CIERRE DEL EJERCICIO ANTERIOR"
        xlHoja1.Cells(9, 21) = "DEPRECIACION DEL EJERCICIO"
        
        xlHoja1.Cells(9, 22) = "DEPRECIACION DEL EJERCICIO RELACIONADA CON LOS RETIROS Y/O BAJAS"
        xlHoja1.Cells(9, 23) = "DEPRECIACION RELACIONADA CON OTROS AJUSTES"
        xlHoja1.Cells(9, 24) = "DEPRECIACION ACUMULADA HISTORICA"
        xlHoja1.Cells(9, 25) = "AJUSTE POR INFLACION DE LA DEPRECIACION"
        xlHoja1.Cells(9, 26) = "DEPRECIACION ACUMULADA AJUSTADA POR INFLACION"
    
    End If
    
    xlHoja1.Range("A3:Z6").Font.Bold = True
    xlHoja1.Range("A8:Z9").Font.Bold = True
    
    xlHoja1.Range("A4:B6").Font.Size = 9
    xlHoja1.Range("A3:Z3").Font.Size = 9
    xlHoja1.Range("A8:Z9").Font.Size = 7
    
    If pbDepre Then '*** PEAC 20110808
    
        xlHoja1.Range("A8:Z9").BorderAround xlContinuous, xlMedium
        xlHoja1.Range("A8:Z9").Borders(xlInsideHorizontal).LineStyle = xlContinuous
        xlHoja1.Range("A8:Z9").Borders(xlInsideVertical).LineStyle = xlContinuous
    
    Else
    
        xlHoja1.Range("A8:P9").BorderAround xlContinuous, xlMedium
        xlHoja1.Range("A8:P9").Borders(xlInsideHorizontal).LineStyle = xlContinuous
        xlHoja1.Range("A8:P9").Borders(xlInsideVertical).LineStyle = xlContinuous
    
    End If
    
    xlHoja1.Range("K2:N2").MergeCells = True
    xlHoja1.Range("K2:L2").Font.Bold = True
     
    xlHoja1.Range("A8:A9").MergeCells = True
    xlHoja1.Range("B8:B9").MergeCells = True
    xlHoja1.Range("C8:F8").MergeCells = True
    xlHoja1.Range("G8:G9").MergeCells = True
    xlHoja1.Range("H8:H9").MergeCells = True
    xlHoja1.Range("I8:I9").MergeCells = True
    xlHoja1.Range("J8:J9").MergeCells = True
    xlHoja1.Range("K8:K9").MergeCells = True
    xlHoja1.Range("L8:L9").MergeCells = True
    xlHoja1.Range("M8:M9").MergeCells = True
    xlHoja1.Range("N8:N9").MergeCells = True
    xlHoja1.Range("O8:O9").MergeCells = True
    xlHoja1.Range("P8:P9").MergeCells = True
    xlHoja1.Range("Q8:R8").MergeCells = True
    xlHoja1.Range("S8:S9").MergeCells = True
    xlHoja1.Range("T8:T9").MergeCells = True
    xlHoja1.Range("U8:U9").MergeCells = True
    
    xlHoja1.Range("V8:V9").MergeCells = True
    xlHoja1.Range("W8:W9").MergeCells = True
    xlHoja1.Range("X8:X9").MergeCells = True
    xlHoja1.Range("Y8:Y9").MergeCells = True
    xlHoja1.Range("Z8:Z9").MergeCells = True
    
    xlHoja1.Range("A4:A4").ColumnWidth = 15
    xlHoja1.Range("B4:B4").ColumnWidth = 15
    xlHoja1.Range("C4:C4").ColumnWidth = 50
    xlHoja1.Range("D4:D4").ColumnWidth = 12
    xlHoja1.Range("E4:E4").ColumnWidth = 12
    xlHoja1.Range("F4:F4").ColumnWidth = 12
    xlHoja1.Range("G4:G4").ColumnWidth = 12
    
    xlHoja1.Range("H4:H4").ColumnWidth = 14
    xlHoja1.Range("I4:I4").ColumnWidth = 11
    xlHoja1.Range("J4:J4").ColumnWidth = 11
    xlHoja1.Range("K4:K4").ColumnWidth = 15
    xlHoja1.Range("L4:L4").ColumnWidth = 15
    
    xlHoja1.Range("M4:M4").ColumnWidth = 9
    xlHoja1.Range("N4:N4").ColumnWidth = 15
    xlHoja1.Range("O4:O4").ColumnWidth = 13
    xlHoja1.Range("P4:P4").ColumnWidth = 15
    
    xlHoja1.Range("Q4:Q4").ColumnWidth = 10
    xlHoja1.Range("R4:R4").ColumnWidth = 15
    xlHoja1.Range("S4:S4").ColumnWidth = 13
    
    xlHoja1.Range("T4:T4").ColumnWidth = 23
    xlHoja1.Range("U4:U4").ColumnWidth = 13
    
    xlHoja1.Range("V4:V4").ColumnWidth = 25
    xlHoja1.Range("W4:W4").ColumnWidth = 16
    xlHoja1.Range("X4:X4").ColumnWidth = 16
    xlHoja1.Range("Y4:Y4").ColumnWidth = 16
    xlHoja1.Range("Z4:Z4").ColumnWidth = 20
    
    xlHoja1.Range("B4:B4").HorizontalAlignment = xlLeft
    xlHoja1.Range("B5:B5").HorizontalAlignment = xlLeft
    xlHoja1.Range("B6:B6").HorizontalAlignment = xlLeft
    
     If pbDepre Then '*** PEAC 20110808
        xlHoja1.Range("A8:Z9").HorizontalAlignment = xlCenter
        xlHoja1.Range("A8:Z9").VerticalAlignment = xlCenter
    Else
        xlHoja1.Range("A8:P9").HorizontalAlignment = xlCenter
        xlHoja1.Range("A8:P9").VerticalAlignment = xlCenter
    End If
    
End Function

Private Sub Form_Load()
    Set oArea = New DActualizaDatosArea
    Me.TxtAgencia.rs = oArea.GetAgencias
    cmbBien.AddItem "TODOS", 0
    cmbBien.Text = "TODOS"
    CargarTipoAF
    txtPeriodo.Text = Format(gdFecSis, "yyyy")
    chkTodos.value = 1
    CargarMes
    Me.DTPicker2.value = CStr(gdFecSis)
End Sub

Private Sub CargarMes()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim oGen As DGeneral
    Set oGen = New DGeneral
    
    Set rs = oGen.GetConstante(1010)
    Me.cmbMes.Clear
    While Not rs.EOF
        cmbMes.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
        If IIf(Len(rs.Fields(1)) = 1, "0" & rs.Fields(1), rs.Fields(1)) = Format(gdFecSis, "MM") Then
            cmbMes.Text = rs.Fields(0) & Space(50) & rs.Fields(1)
        End If
        rs.MoveNext
    Wend
End Sub

Private Sub CargarTipoAF()
    Dim i As Integer
    Dim rsDatos As ADODB.Recordset
    Set oInventario = New NInvActivoFijo
    Set rsDatos = oInventario.ObtenerTipoAF
    
    For i = 0 To rsDatos.RecordCount - 1
        cmbTipo.AddItem rsDatos.Fields(0) & Space(50) & rsDatos.Fields(1)
        rsDatos.MoveNext
    Next i
    cmbTipo.AddItem "TODOS", 0
    cmbTipo.Text = "TODOS"
    Set rsDatos = Nothing
End Sub

Private Sub chkTodos_Click()
    If Me.chkTodos.value = 1 Then
        Me.TxtAgencia.Text = ""
        Me.lblAgencia.Caption = ""
    End If
End Sub

Private Sub txtAgencia_EmiteDatos()
    lblAgencia.Caption = TxtAgencia.psDescripcion
    chkTodos.value = 0
End Sub
