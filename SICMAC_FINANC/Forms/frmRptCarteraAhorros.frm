VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRptCarteraAhorros 
   Caption         =   "Reportes Balance Planeamiento: Cartera de Ahorros"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRep 
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Frame Frame2 
         Height          =   525
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   3120
         Begin VB.OptionButton OptAnalista 
            Caption         =   "&Ninguno"
            Height          =   210
            Index           =   1
            Left            =   1620
            TabIndex        =   19
            Top             =   195
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton OptAnalista 
            Caption         =   "&Todos"
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   195
            Width           =   915
         End
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Agencias a escoger "
         Height          =   2895
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   3165
         Begin VB.ListBox lstAge 
            Height          =   2535
            ItemData        =   "frmRptCarteraAhorros.frx":0000
            Left            =   120
            List            =   "frmRptCarteraAhorros.frx":0002
            Style           =   1  'Checkbox
            TabIndex        =   16
            Top             =   240
            Width           =   2925
         End
      End
      Begin VB.Frame fraMes 
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
         Height          =   1965
         Left            =   3480
         TabIndex        =   3
         Top             =   120
         Width           =   3735
         Begin VB.ComboBox cboMesHasta 
            Height          =   315
            ItemData        =   "frmRptCarteraAhorros.frx":0004
            Left            =   2040
            List            =   "frmRptCarteraAhorros.frx":002C
            TabIndex        =   14
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txtAnioHasta 
            Alignment       =   1  'Right Justify
            Height          =   280
            Left            =   600
            TabIndex        =   12
            Top             =   1440
            Width           =   780
         End
         Begin VB.TextBox txtAnio 
            Alignment       =   1  'Right Justify
            Height          =   280
            Left            =   600
            MaxLength       =   4
            TabIndex        =   5
            Top             =   660
            Width           =   780
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            ItemData        =   "frmRptCarteraAhorros.frx":0094
            Left            =   2040
            List            =   "frmRptCarteraAhorros.frx":00BC
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   660
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Mes"
            Height          =   255
            Left            =   1560
            TabIndex        =   13
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Año"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta:"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Desde:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Año :"
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mes :"
            Height          =   195
            Left            =   1560
            TabIndex        =   6
            Top             =   720
            Width           =   390
         End
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   345
         Left            =   6000
         TabIndex        =   2
         Top             =   2640
         Width           =   1155
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "&Generar"
         Height          =   345
         Left            =   4800
         TabIndex        =   1
         Top             =   2640
         Width           =   1155
      End
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   2160
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "frmRptCarteraAhorros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmRptCarteraAhorros
'** Descripción : Generación de Reportes Cartera de Ahorros segun ERS165-2013
'** Creación : FRHU, 20140120 09:00:00 AM
'********************************************************************
Private Sub cmdGenerar_Click()
    If Me.txtAnio.Text = "" Or Me.txtAnioHasta.Text = "" Or cboMes.ListIndex = -1 Or Me.cboMesHasta.ListIndex = -1 Then
        MsgBox "Debe llenar todos los campos del Periodo", vbInformation
        Exit Sub
    Else
        Call GenerarRptCarteraAhorros
    End If
End Sub
Private Sub GenerarRptCarteraAhorros()
    Dim oAgencia As New DAgencia
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet, xlsHoja1 As Excel.Worksheet, xlsHoja2 As Excel.Worksheet
    Dim rsAgencia As New ADODB.Recordset
    Dim ldFecha As Date
    Dim lsArchivo As String
    Dim i As Integer
    Dim CodAgencia As String
    Dim lnCuenta As Integer
    Dim lnValor As Integer
    lnValor = 0
    On Error GoTo ErrGenerarRptCarteraAhorros
    
    lsArchivo = "\spooler\ReporteCarteraAhorros" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    Set xlsLibro = xlsAplicacion.Workbooks.Add
    
    lnCuenta = Me.lstAge.ListCount
    PB1.Min = 0
    PB1.Max = lnCuenta
    PB1.value = 0
    PB1.Visible = True
    Me.MousePointer = vbHourglass
    
    For i = 0 To Me.lstAge.ListCount - 1
        PB1.value = i
        If Me.lstAge.Selected(i) Then
            lnValor = 1
            CodAgencia = Right(Me.lstAge.List(i), 2)
            Set xlsHoja2 = xlsLibro.Worksheets.Add
            Set xlsHoja1 = xlsLibro.Worksheets.Add
            Set xlsHoja = xlsLibro.Worksheets.Add
            xlsHoja.Name = "DATOS DEP-Ag" & CodAgencia
            xlsHoja1.Name = "DEP x Producto Ag" & CodAgencia
            xlsHoja2.Name = "DEP TOTALES Ag" & CodAgencia
            'Hoja 00 Y 01
            Call generaHojaExcelAgencia_RptCarteraAhorros(xlsHoja, xlsHoja1, CodAgencia, CStr(Me.txtAnio.Text), CStr(cboMes.ListIndex + 1), CStr(Me.txtAnioHasta.Text), CStr(Me.cboMesHasta.ListIndex + 1))
            'Hoja 02
            PB1.value = i + 1
            Call generaHojaExcelAgencia_RptDepositosTotales(xlsHoja2, CodAgencia, "DEP TOTALES Ag" & CodAgencia, CStr(Me.txtAnio.Text), CStr(cboMes.ListIndex + 1), CStr(Me.txtAnioHasta.Text), CStr(Me.cboMesHasta.ListIndex + 1))
        Else
            PB1.value = i + 1
        End If
    Next i
    
    If lnValor = 0 Then
        Me.MousePointer = vbDefault
        PB1.Visible = False
        MsgBox "Seleccione por lo menos una Agencia", vbInformation, "Sistema"
        Exit Sub
    End If
    
    xlsHoja2.SaveAs App.path & lsArchivo
    xlsHoja1.SaveAs App.path & lsArchivo
    xlsHoja.SaveAs App.path & lsArchivo
    
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    xlsAplicacion.UserControl = True
    
    Set oAgencia = Nothing
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlsHoja = Nothing
    Set xlsHoja1 = Nothing
    Set xlsHoja2 = Nothing
    PB1.Visible = False
    Me.MousePointer = vbDefault
    Exit Sub
ErrGenerarRptCarteraAhorros:
    PB1.Visible = False
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, "Aviso"
    Exit Sub
End Sub
Private Sub generaHojaExcelAgencia_RptCarteraAhorros(ByRef xlsHoja As Worksheet, ByRef xlsHoja1 As Worksheet, ByVal psAgeCod As String, ByVal psAnio As String, ByVal psMes As String, ByVal psAnioHasta As String, ByVal psMesHasta As String)
    Dim oBalance As New DbalanceCont
    Dim oCaja As New DCajaGeneral
    Dim rsAhorros As New ADODB.Recordset
    Dim oAgencia As New DAgencia
    Dim cAgencia As String
    Dim lnLineaActual As Integer, lnLineaHoja1 As Integer
    Dim ldFecha As Date
    Dim lnSaldoEval As Currency
    Dim lnVarEval As Double
    Dim lnFilaInicio As Integer
    Dim lnPrimeraFila As Integer
    lnPrimeraFila = 1
    
    ldFecha = CDate("31/12/" & Format(Val(txtAnioHasta.Text) - 1, "0000"))
    
    cAgencia = UCase(Trim(oAgencia.GetAgencias(psAgeCod)))
    xlsHoja.Range("B1", "I1").MergeCells = True
    xlsHoja.Range("B1", "I1").HorizontalAlignment = xlCenter
    xlsHoja.Range("B1:I1").Interior.Color = RGB(255, 255, 0)
    xlsHoja.Range("B1:I1").Font.Size = 15
    xlsHoja.Range("B1", "I1").Font.Bold = True
    xlsHoja.Range("B1") = "EVOLUCION DE LOS DEPOSITOS POR PRODUCTOS - " & cAgencia
    
    xlsHoja.Cells.Font.Name = "Calibri"
    'xlsHoja.Cells.Font.Size = 10
    xlsHoja.Range("A:A").ColumnWidth = 4
    xlsHoja.Range("B:B").ColumnWidth = 11
    xlsHoja.Range("C:C").ColumnWidth = 11
    xlsHoja.Range("D:D").ColumnWidth = 12
    xlsHoja.Range("E:E").ColumnWidth = 12
    xlsHoja.Range("F:F").ColumnWidth = 12
    xlsHoja.Range("G:G").ColumnWidth = 16
    xlsHoja.Range("H:H").ColumnWidth = 16
    xlsHoja.Range("I:I").ColumnWidth = 16
    'xlsHoja.Range("B13").RowHeight = 3
    'xlsHoja.Range("B1") = "EVOLUCION DE LOS DEPOSITOS POR PRODUCTOS - " & cAgencia
    xlsHoja.Range("B3", "I3").MergeCells = True
    xlsHoja.Range("B3", "I3").HorizontalAlignment = xlCenter
    xlsHoja.Range("B3", "I4").Font.Bold = True
    xlsHoja.Range("B3:I3").Font.Size = 14
    xlsHoja.Range("B3") = "AHORRO CORRIENTE"
    xlsHoja.Range("B4", "C4").MergeCells = True
    xlsHoja.Range("B4") = cAgencia
    xlsHoja.Range("D4", "H4").MergeCells = True
    xlsHoja.Range("D4", "H4").HorizontalAlignment = xlCenter
    xlsHoja.Range("D4") = " ( en nuevo soles ) "
    xlsHoja.Range("I4", "I4").HorizontalAlignment = xlRight
    xlsHoja.Range("I4") = "Año " & psAnioHasta
    'CABECERA HOJA000
    xlsHoja.Range("B5", "I6").HorizontalAlignment = xlCenter
    xlsHoja.Range("B5", "I6").Font.Bold = True
    xlsHoja.Range("B5:I6").Interior.Color = RGB(199, 199, 199)
    xlsHoja.Range("B5", "B6").MergeCells = True
    xlsHoja.Range("B5") = "Periodo"
    xlsHoja.Range("C5", "C6").MergeCells = True
    xlsHoja.Range("C5") = "Nº Ctas."
    xlsHoja.Range("D5", "F5").MergeCells = True
    xlsHoja.Range("D5") = "Importe S/."
    xlsHoja.Range("D6") = "Total"
    xlsHoja.Range("E6") = "MN"
    xlsHoja.Range("F6") = "ME"
    xlsHoja.Range("G5", "G6").MergeCells = True
    xlsHoja.Range("G5") = "Crecimiento Anual / Mensual"
    xlsHoja.Range("H5", "H6").MergeCells = True
    xlsHoja.Range("H5") = "Var. mes %"
    xlsHoja.Range("I5", "I6").MergeCells = True
    xlsHoja.Range("I5") = "Var. acum %"
    'CABECERA HOJA001
    xlsHoja1.Range("B3", "J3").Font.Bold = True
    xlsHoja1.Range("B3:J3").Interior.Color = RGB(255, 255, 0)
    xlsHoja1.Range("B3", "J3").MergeCells = True
    xlsHoja1.Range("B3:J3").Font.Size = 15
    xlsHoja1.Range("B3", "J3").HorizontalAlignment = xlCenter
    xlsHoja1.Range("B3") = "Evolucion y Participación de los DEPOSITOS"
    xlsHoja1.Range("B4", "C4").MergeCells = True
    xlsHoja1.Range("B4") = cAgencia
    xlsHoja1.Range("D4", "H4").MergeCells = True
    xlsHoja1.Range("D4", "H4").HorizontalAlignment = xlCenter
    xlsHoja1.Range("D4") = " ( en nuevo soles ) "
    xlsHoja1.Range("I4", "J4").MergeCells = True
    xlsHoja1.Range("I4", "J4").HorizontalAlignment = xlRight
    xlsHoja1.Range("I4") = "Año " & psAnioHasta
    xlsHoja1.Range("B4", "J6").Font.Bold = True
    xlsHoja1.Range("B5:J6").Interior.Color = RGB(204, 255, 204)
    xlsHoja1.Range("B5") = "PRODUCTO"
    xlsHoja1.Range("B6") = "Meses"
    xlsHoja1.Range("C5", "D5").MergeCells = True
    xlsHoja1.Range("C5") = "AHORRO CORRIENTE"
    xlsHoja1.Range("C6") = "Saldos"
    xlsHoja1.Range("D6") = "Var. %"
    xlsHoja1.Range("E5", "F5").MergeCells = True
    xlsHoja1.Range("E5") = "PLAZO FIJO"
    xlsHoja1.Range("E6") = "Saldos"
    xlsHoja1.Range("F6") = "Var. %"
    xlsHoja1.Range("G5", "H5").MergeCells = True
    xlsHoja1.Range("G5") = "CTS"
    xlsHoja1.Range("G6") = "Saldos"
    xlsHoja1.Range("H6") = "Var. %"
    xlsHoja1.Range("I5", "J5").MergeCells = True
    xlsHoja1.Range("I5") = "TOTALES"
    xlsHoja1.Range("I6") = "Saldos"
    xlsHoja1.Range("J6") = "Var. %"
    
    xlsHoja.Range("B" & Trim(Str(5)) & ":" & "I" & Trim(Str(5))).Borders.LineStyle = 1
    xlsHoja1.Range("B" & Trim(Str(5)) & ":" & "J" & Trim(Str(5))).Borders.LineStyle = 1 'HOJA001
    lnLineaActual = 6
    lnLineaHoja1 = 6 'HOJA001
    lnFilaInicio = lnLineaActual + 1
    '*********** AHORRO CORRIENTE
    Set rsAhorros = oCaja.RecuperaRptAhorros(psAnio, psMes, psAnioHasta, psMesHasta, psAgeCod)
    Do While Not rsAhorros.EOF
        If lnPrimeraFila = 1 Then
            lnPrimeraFila = lnPrimeraFila + 1
            rsAhorros.MoveNext
        Else
            xlsHoja.Range("B" & Trim(Str(lnLineaActual)) & ":" & "I" & Trim(Str(lnLineaActual))).Borders.LineStyle = 1
            xlsHoja1.Range("B" & Trim(Str(lnLineaHoja1)) & ":" & "J" & Trim(Str(lnLineaHoja1))).Borders.LineStyle = 1 'HOJA001
            lnLineaHoja1 = lnLineaHoja1 + 1
            lnLineaActual = lnLineaActual + 1
            xlsHoja.Cells(lnLineaActual, 2).Formula = Format(rsAhorros!cPeriodo, "dd/mm/yyyy")
            xlsHoja.Cells(lnLineaActual, 3).Formula = rsAhorros!nNCuentas
            xlsHoja.Cells(lnLineaActual, 4).Formula = rsAhorros!mImporteTotal
            xlsHoja.Cells(lnLineaActual, 5).Formula = rsAhorros!mImporteMN
            xlsHoja.Cells(lnLineaActual, 6).Formula = rsAhorros!mImporteME
            xlsHoja.Cells(lnLineaActual, 7).Formula = rsAhorros!mCrecimiento
            xlsHoja.Cells(lnLineaActual, 8).Formula = rsAhorros!fVarMensual / 100
            xlsHoja.Cells(lnLineaActual, 9).Formula = rsAhorros!fVarAcumulada / 100
            'HOJA002
            xlsHoja1.Cells(lnLineaHoja1, 2).Formula = Format(rsAhorros!cPeriodo, "dd/mm/yyyy")
            xlsHoja1.Cells(lnLineaHoja1, 3).Formula = rsAhorros!mImporteTotal
            xlsHoja1.Cells(lnLineaHoja1, 4).Formula = rsAhorros!fVarMensual / 100
            If ldFecha = CDate(rsAhorros!cPeriodo) Then
                lnSaldoEval = rsAhorros!mImporteTotal
            End If
            rsAhorros.MoveNext
        End If
    Loop
    lnPrimeraFila = 1
    xlsHoja.Range("D" & lnFilaInicio & ":G" & lnLineaActual).NumberFormat = "#,##0"
    xlsHoja.Range("H" & lnFilaInicio & ":I" & lnLineaActual).NumberFormat = "0.00%"
    xlsHoja.Range("B" & Trim(Str(lnLineaActual)) & ":" & "I" & Trim(Str(lnLineaActual))).Borders.LineStyle = 1
    'VAR MENSUAL HOJA001
    xlsHoja1.Range("C" & 7 & ":C" & lnLineaHoja1).NumberFormat = "#,##0"
    xlsHoja1.Range("D" & 7 & ":D" & lnLineaHoja1).NumberFormat = "0.00%"
    xlsHoja1.Cells(lnLineaHoja1 + 2, 3).Formula = xlsHoja1.Cells(lnLineaHoja1, 3).Formula - xlsHoja1.Cells(lnLineaHoja1 - 1, 3).Formula
    'VAR ACUMULADA HOJA001
    xlsHoja1.Cells(lnLineaHoja1 + 3, 3).Formula = xlsHoja1.Cells(lnLineaHoja1, 3).Formula - lnSaldoEval
    xlsHoja1.Cells(lnLineaHoja1 + 3, 4).Formula = xlsHoja1.Cells(lnLineaHoja1, 3).Formula / lnSaldoEval - 1
    
    lnLineaHoja1 = 6 'HOJA001
    '*********** DEPOSITO PLAZO FIJO
    lnLineaActual = lnLineaActual + 2
    xlsHoja.Range("B" & CStr(lnLineaActual + 1), "I" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("B" & CStr(lnLineaActual + 1), "I" & CStr(lnLineaActual + 1)).HorizontalAlignment = xlCenter
    xlsHoja.Range("B" & CStr(lnLineaActual + 1), "I" & CStr(lnLineaActual + 2)).Font.Bold = True
    xlsHoja.Range("B" & CStr(lnLineaActual + 1), "I" & CStr(lnLineaActual + 1)).Font.Size = 14
    xlsHoja.Range("B" & CStr(lnLineaActual + 1)) = "PLAZO FIJO"
    xlsHoja.Range("B" & CStr(lnLineaActual + 2), "C" & CStr(lnLineaActual + 2)).MergeCells = True
    xlsHoja.Range("B" & CStr(lnLineaActual + 2)) = cAgencia
    xlsHoja.Range("D" & CStr(lnLineaActual + 2), "H" & CStr(lnLineaActual + 2)).MergeCells = True
    xlsHoja.Range("D" & CStr(lnLineaActual + 2), "H" & CStr(lnLineaActual + 2)).HorizontalAlignment = xlCenter
    xlsHoja.Range("D" & CStr(lnLineaActual + 2)) = " ( en nuevo soles ) "
    xlsHoja.Range("I" & CStr(lnLineaActual + 2), "I" & CStr(lnLineaActual + 2)).HorizontalAlignment = xlRight
    xlsHoja.Range("I" & CStr(lnLineaActual + 2)) = "Año " & psAnioHasta
    lnLineaActual = lnLineaActual + 3
    'CABECERA
    xlsHoja.Range("B" & CStr(lnLineaActual), "I" & CStr(lnLineaActual + 1)).HorizontalAlignment = xlCenter
    xlsHoja.Range("B" & CStr(lnLineaActual), "I" & CStr(lnLineaActual + 1)).Font.Bold = True
    xlsHoja.Range("B" & CStr(lnLineaActual), "I" & CStr(lnLineaActual + 1)).Interior.Color = RGB(199, 199, 199)
    xlsHoja.Range("B" & CStr(lnLineaActual), "B" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("B" & CStr(lnLineaActual)) = "Periodo"
    xlsHoja.Range("C" & CStr(lnLineaActual), "C" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("C" & CStr(lnLineaActual)) = "Nº Ctas."
    xlsHoja.Range("D" & CStr(lnLineaActual), "F" & CStr(lnLineaActual)).MergeCells = True
    xlsHoja.Range("D" & CStr(lnLineaActual)) = "Importe S/."
    xlsHoja.Range("D" & CStr(lnLineaActual + 1)) = "Total"
    xlsHoja.Range("E" & CStr(lnLineaActual + 1)) = "MN"
    xlsHoja.Range("F" & CStr(lnLineaActual + 1)) = "ME"
    xlsHoja.Range("G" & CStr(lnLineaActual), "G" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("G" & CStr(lnLineaActual)) = "Crecimiento Anual / Mensual"
    xlsHoja.Range("H" & CStr(lnLineaActual), "H" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("H" & CStr(lnLineaActual)) = "Var. mes %"
    xlsHoja.Range("I" & CStr(lnLineaActual), "I" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("I" & CStr(lnLineaActual)) = "Var. acum %"
    xlsHoja.Range("B" & Trim(Str(lnLineaActual)) & ":" & "I" & Trim(Str(lnLineaActual))).Borders.LineStyle = 1
    'xlsHoja.Range("B" & Trim(Str(lnLineaActual + 1)) & ":" & "I" & Trim(Str(lnLineaActual + 1))).Borders.LineStyle = 1
    lnLineaActual = lnLineaActual + 1
    lnFilaInicio = lnLineaActual + 1
    Set rsAhorros = oCaja.RecuperaRptDepositoPlazoFijo(psAnio, psMes, psAnioHasta, psMesHasta, psAgeCod)
    Do While Not rsAhorros.EOF
        If lnPrimeraFila = 1 Then
            lnPrimeraFila = lnPrimeraFila + 1
            rsAhorros.MoveNext
        Else
            xlsHoja.Range("B" & Trim(Str(lnLineaActual)) & ":" & "I" & Trim(Str(lnLineaActual))).Borders.LineStyle = 1
            lnLineaActual = lnLineaActual + 1
            lnLineaHoja1 = lnLineaHoja1 + 1 'HOJA001
            xlsHoja.Cells(lnLineaActual, 2).Formula = Format(rsAhorros!cPeriodo, "dd/mm/yyyy")
            xlsHoja.Cells(lnLineaActual, 3).Formula = rsAhorros!nNCuentas
            xlsHoja.Cells(lnLineaActual, 4).Formula = rsAhorros!mImporteTotal
            xlsHoja.Cells(lnLineaActual, 5).Formula = rsAhorros!mImporteMN
            xlsHoja.Cells(lnLineaActual, 6).Formula = rsAhorros!mImporteME
            xlsHoja.Cells(lnLineaActual, 7).Formula = rsAhorros!mCrecimiento
            xlsHoja.Cells(lnLineaActual, 8).Formula = rsAhorros!fVarMensual / 100
            xlsHoja.Cells(lnLineaActual, 9).Formula = rsAhorros!fVarAcumulada / 100
            'HOJA002
            xlsHoja1.Cells(lnLineaHoja1, 5).Formula = rsAhorros!mImporteTotal
            xlsHoja1.Cells(lnLineaHoja1, 6).Formula = rsAhorros!fVarMensual / 100
            If ldFecha = CDate(rsAhorros!cPeriodo) Then
                lnSaldoEval = rsAhorros!mImporteTotal
            End If
            rsAhorros.MoveNext
        End If
    Loop
    lnPrimeraFila = 1
    xlsHoja.Range("D" & lnFilaInicio & ":G" & lnLineaActual).NumberFormat = "#,##0"
    xlsHoja.Range("H" & lnFilaInicio & ":I" & lnLineaActual).NumberFormat = "0.00%"
    xlsHoja.Range("B" & Trim(Str(lnLineaActual)) & ":" & "I" & Trim(Str(lnLineaActual))).Borders.LineStyle = 1
   'VAR MENSUAL HOJA001
    xlsHoja1.Range("E" & 7 & ":E" & lnLineaHoja1).NumberFormat = "#,##0"
    xlsHoja1.Range("F" & 7 & ":F" & lnLineaHoja1).NumberFormat = "0.00%"
    xlsHoja1.Cells(lnLineaHoja1 + 2, 5).Formula = xlsHoja1.Cells(lnLineaHoja1, 5).Formula - xlsHoja1.Cells(lnLineaHoja1 - 1, 5).Formula
    'VAR ACUMULADA HOJA001
    xlsHoja1.Cells(lnLineaHoja1 + 3, 5).Formula = xlsHoja1.Cells(lnLineaHoja1, 5).Formula - lnSaldoEval
    xlsHoja1.Cells(lnLineaHoja1 + 3, 6).Formula = xlsHoja1.Cells(lnLineaHoja1, 5).Formula / lnSaldoEval - 1
    lnLineaHoja1 = 6 'HOJA001
    '*********** DEPOSITO CTS
    lnLineaActual = lnLineaActual + 2
    xlsHoja.Range("B" & CStr(lnLineaActual + 1), "I" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("B" & CStr(lnLineaActual + 1), "I" & CStr(lnLineaActual + 1)).HorizontalAlignment = xlCenter
    xlsHoja.Range("B" & CStr(lnLineaActual + 1), "I" & CStr(lnLineaActual + 2)).Font.Bold = True
    xlsHoja.Range("B" & CStr(lnLineaActual + 1), "I" & CStr(lnLineaActual + 1)).Font.Size = 14
    xlsHoja.Range("B" & CStr(lnLineaActual + 1)) = "DEPOSITO CTS"
    xlsHoja.Range("B" & CStr(lnLineaActual + 2), "C" & CStr(lnLineaActual + 2)).MergeCells = True
    xlsHoja.Range("B" & CStr(lnLineaActual + 2)) = cAgencia
    xlsHoja.Range("D" & CStr(lnLineaActual + 2), "H" & CStr(lnLineaActual + 2)).MergeCells = True
    xlsHoja.Range("D" & CStr(lnLineaActual + 2), "H" & CStr(lnLineaActual + 2)).HorizontalAlignment = xlCenter
    xlsHoja.Range("D" & CStr(lnLineaActual + 2)) = " ( en nuevo soles ) "
    xlsHoja.Range("I" & CStr(lnLineaActual + 2), "I" & CStr(lnLineaActual + 2)).HorizontalAlignment = xlRight
    xlsHoja.Range("I" & CStr(lnLineaActual + 2)) = "Año " & psAnioHasta
    lnLineaActual = lnLineaActual + 3
    'CABECERA
    xlsHoja.Range("B" & CStr(lnLineaActual), "I" & CStr(lnLineaActual + 1)).HorizontalAlignment = xlCenter
    xlsHoja.Range("B" & CStr(lnLineaActual), "I" & CStr(lnLineaActual + 1)).Font.Bold = True
    xlsHoja.Range("B" & CStr(lnLineaActual), "I" & CStr(lnLineaActual + 1)).Interior.Color = RGB(199, 199, 199)
    xlsHoja.Range("B" & CStr(lnLineaActual), "B" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("B" & CStr(lnLineaActual)) = "Periodo"
    xlsHoja.Range("C" & CStr(lnLineaActual), "C" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("C" & CStr(lnLineaActual)) = "Nº Ctas."
    xlsHoja.Range("D" & CStr(lnLineaActual), "F" & CStr(lnLineaActual)).MergeCells = True
    xlsHoja.Range("D" & CStr(lnLineaActual)) = "Importe S/."
    xlsHoja.Range("D" & CStr(lnLineaActual + 1)) = "Total"
    xlsHoja.Range("E" & CStr(lnLineaActual + 1)) = "MN"
    xlsHoja.Range("F" & CStr(lnLineaActual + 1)) = "ME"
    xlsHoja.Range("G" & CStr(lnLineaActual), "G" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("G" & CStr(lnLineaActual)) = "Crecimiento Anual / Mensual"
    xlsHoja.Range("H" & CStr(lnLineaActual), "H" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("H" & CStr(lnLineaActual)) = "Var. mes %"
    xlsHoja.Range("I" & CStr(lnLineaActual), "I" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("I" & CStr(lnLineaActual)) = "Var. acum %"
    xlsHoja.Range("B" & Trim(Str(lnLineaActual)) & ":" & "I" & Trim(Str(lnLineaActual))).Borders.LineStyle = 1
    'xlsHoja.Range("B" & Trim(Str(lnLineaActual + 1)) & ":" & "I" & Trim(Str(lnLineaActual + 1))).Borders.LineStyle = 1
    lnLineaActual = lnLineaActual + 1
    lnFilaInicio = lnLineaActual + 1
    Set rsAhorros = oCaja.RecuperaRptDepositoCTS(psAnio, psMes, psAnioHasta, psMesHasta, psAgeCod)
    Do While Not rsAhorros.EOF
        If lnPrimeraFila = 1 Then
            lnPrimeraFila = lnPrimeraFila + 1
            rsAhorros.MoveNext
        Else
            xlsHoja.Range("B" & Trim(Str(lnLineaActual)) & ":" & "I" & Trim(Str(lnLineaActual))).Borders.LineStyle = 1
            lnLineaActual = lnLineaActual + 1
            lnLineaHoja1 = lnLineaHoja1 + 1 'HOJA001
            xlsHoja.Cells(lnLineaActual, 2).Formula = Format(rsAhorros!cPeriodo, "dd/mm/yyyy")
            xlsHoja.Cells(lnLineaActual, 3).Formula = rsAhorros!nNCuentas
            xlsHoja.Cells(lnLineaActual, 4).Formula = rsAhorros!mImporteTotal
            xlsHoja.Cells(lnLineaActual, 5).Formula = rsAhorros!mImporteMN
            xlsHoja.Cells(lnLineaActual, 6).Formula = rsAhorros!mImporteME
            xlsHoja.Cells(lnLineaActual, 7).Formula = rsAhorros!mCrecimiento
            xlsHoja.Cells(lnLineaActual, 8).Formula = rsAhorros!fVarMensual / 100
            xlsHoja.Cells(lnLineaActual, 9).Formula = rsAhorros!fVarAcumulada / 100
            'HOJA002
            xlsHoja1.Cells(lnLineaHoja1, 7).Formula = rsAhorros!mImporteTotal
            xlsHoja1.Cells(lnLineaHoja1, 8).Formula = rsAhorros!fVarMensual / 100
            If ldFecha = CDate(rsAhorros!cPeriodo) Then
                lnSaldoEval = rsAhorros!mImporteTotal
            End If
            rsAhorros.MoveNext
        End If
    Loop
    lnPrimeraFila = 1
    xlsHoja.Range("D" & lnFilaInicio & ":G" & lnLineaActual).NumberFormat = "#,##0"
    xlsHoja.Range("H" & lnFilaInicio & ":I" & lnLineaActual).NumberFormat = "0.00%"
    xlsHoja.Range("B" & Trim(Str(lnLineaActual)) & ":" & "I" & Trim(Str(lnLineaActual))).Borders.LineStyle = 1
    'VAR MENSUAL HOJA001
    xlsHoja1.Cells(lnLineaHoja1 + 2, 7).Formula = xlsHoja1.Cells(lnLineaHoja1, 7).Formula - xlsHoja1.Cells(lnLineaHoja1 - 1, 7).Formula
    'VAR ACUMULADA HOJA001
    xlsHoja1.Range("G" & 7 & ":G" & lnLineaHoja1).NumberFormat = "#,##0"
    xlsHoja1.Range("H" & 7 & ":H" & lnLineaHoja1).NumberFormat = "0.00%"
    xlsHoja1.Cells(lnLineaHoja1 + 3, 7).Formula = xlsHoja1.Cells(lnLineaHoja1, 7).Formula - lnSaldoEval
    xlsHoja1.Cells(lnLineaHoja1 + 3, 8).Formula = xlsHoja1.Cells(lnLineaHoja1, 7).Formula / lnSaldoEval - 1
    lnLineaHoja1 = 6 'HOJA001
    '*********** DEPOSITO TOTAL
    lnLineaActual = lnLineaActual + 2
    xlsHoja.Range("B" & CStr(lnLineaActual + 1), "I" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("B" & CStr(lnLineaActual + 1), "I" & CStr(lnLineaActual + 1)).HorizontalAlignment = xlCenter
    xlsHoja.Range("B" & CStr(lnLineaActual + 1), "I" & CStr(lnLineaActual + 2)).Font.Bold = True
    xlsHoja.Range("B" & CStr(lnLineaActual + 1), "I" & CStr(lnLineaActual + 1)).Font.Size = 14
    xlsHoja.Range("B" & CStr(lnLineaActual + 1)) = "DEPOSITO TOTAL"
    xlsHoja.Range("B" & CStr(lnLineaActual + 2), "C" & CStr(lnLineaActual + 2)).MergeCells = True
    xlsHoja.Range("B" & CStr(lnLineaActual + 2)) = cAgencia
    xlsHoja.Range("D" & CStr(lnLineaActual + 2), "H" & CStr(lnLineaActual + 2)).MergeCells = True
    xlsHoja.Range("D" & CStr(lnLineaActual + 2), "H" & CStr(lnLineaActual + 2)).HorizontalAlignment = xlCenter
    xlsHoja.Range("D" & CStr(lnLineaActual + 2)) = " ( en nuevo soles ) "
    xlsHoja.Range("I" & CStr(lnLineaActual + 2), "I" & CStr(lnLineaActual + 2)).HorizontalAlignment = xlRight
    xlsHoja.Range("I" & CStr(lnLineaActual + 2)) = "Año " & psAnioHasta
    
    lnLineaActual = lnLineaActual + 3
    'CABECERA
    xlsHoja.Range("B" & CStr(lnLineaActual), "I" & CStr(lnLineaActual + 1)).HorizontalAlignment = xlCenter
    xlsHoja.Range("B" & CStr(lnLineaActual), "I" & CStr(lnLineaActual + 1)).Font.Bold = True
    xlsHoja.Range("B" & CStr(lnLineaActual), "I" & CStr(lnLineaActual + 1)).Interior.Color = RGB(199, 199, 199)
    xlsHoja.Range("B" & CStr(lnLineaActual), "B" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("B" & CStr(lnLineaActual)) = "Periodo"
    xlsHoja.Range("C" & CStr(lnLineaActual), "C" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("C" & CStr(lnLineaActual)) = "Nº Ctas."
    xlsHoja.Range("D" & CStr(lnLineaActual), "F" & CStr(lnLineaActual)).MergeCells = True
    xlsHoja.Range("D" & CStr(lnLineaActual)) = "Importe S/."
    xlsHoja.Range("D" & CStr(lnLineaActual + 1)) = "Total"
    xlsHoja.Range("E" & CStr(lnLineaActual + 1)) = "MN"
    xlsHoja.Range("F" & CStr(lnLineaActual + 1)) = "ME"
    xlsHoja.Range("G" & CStr(lnLineaActual), "G" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("G" & CStr(lnLineaActual)) = "Crecimiento Anual / Mensual"
    xlsHoja.Range("H" & CStr(lnLineaActual), "H" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("H" & CStr(lnLineaActual)) = "Var. mes %"
    xlsHoja.Range("I" & CStr(lnLineaActual), "I" & CStr(lnLineaActual + 1)).MergeCells = True
    xlsHoja.Range("I" & CStr(lnLineaActual)) = "Var. acum %"
    xlsHoja.Range("B" & Trim(Str(lnLineaActual)) & ":" & "I" & Trim(Str(lnLineaActual))).Borders.LineStyle = 1
    lnLineaActual = lnLineaActual + 1
    lnFilaInicio = lnLineaActual + 1
    Set rsAhorros = oCaja.RecuperaRptDepositoTOTAL(psAnio, psMes, psAnioHasta, psMesHasta, psAgeCod)
    Do While Not rsAhorros.EOF
        If lnPrimeraFila = 1 Then
            lnPrimeraFila = lnPrimeraFila + 1
            rsAhorros.MoveNext
        Else
            xlsHoja.Range("B" & Trim(Str(lnLineaActual)) & ":" & "I" & Trim(Str(lnLineaActual))).Borders.LineStyle = 1
            lnLineaActual = lnLineaActual + 1
            lnLineaHoja1 = lnLineaHoja1 + 1 'HOJA001
            xlsHoja.Cells(lnLineaActual, 2).Formula = Format(rsAhorros!cPeriodo, "dd/mm/yyyy")
            xlsHoja.Cells(lnLineaActual, 3).Formula = rsAhorros!nNCuentas
            xlsHoja.Cells(lnLineaActual, 4).Formula = rsAhorros!mImporteTotal
            xlsHoja.Cells(lnLineaActual, 5).Formula = rsAhorros!mImporteMN
            xlsHoja.Cells(lnLineaActual, 6).Formula = rsAhorros!mImporteME
            xlsHoja.Cells(lnLineaActual, 7).Formula = rsAhorros!mCrecimiento
            xlsHoja.Cells(lnLineaActual, 8).Formula = rsAhorros!fVarMensual / 100
            xlsHoja.Cells(lnLineaActual, 9).Formula = rsAhorros!fVarAcumulada / 100
            'HOJA002
            xlsHoja1.Cells(lnLineaHoja1, 9).Formula = rsAhorros!mImporteTotal
            xlsHoja1.Cells(lnLineaHoja1, 10).Formula = rsAhorros!fVarMensual / 100
            If ldFecha = CDate(rsAhorros!cPeriodo) Then
                lnSaldoEval = rsAhorros!mImporteTotal
            End If
            rsAhorros.MoveNext
        End If
    Loop
    xlsHoja.Range("D" & lnFilaInicio & ":G" & lnLineaActual).NumberFormat = "#,##0"
    xlsHoja.Range("H" & lnFilaInicio & ":I" & lnLineaActual).NumberFormat = "0.00%"
    xlsHoja.Range("B" & Trim(Str(lnLineaActual)) & ":" & "I" & Trim(Str(lnLineaActual))).Borders.LineStyle = 1
    xlsHoja1.Range("B" & Trim(Str(lnLineaHoja1)) & ":" & "J" & Trim(Str(lnLineaHoja1))).Borders.LineStyle = 1 'HOJA001
    
    'VAR MENSUAL HOJA001
    xlsHoja1.Range("I" & 7 & ":I" & lnLineaHoja1).NumberFormat = "#,##0"
    xlsHoja1.Range("J" & 7 & ":J" & lnLineaHoja1).NumberFormat = "0.00%"
    xlsHoja1.Cells(lnLineaHoja1 + 2, 9).Formula = xlsHoja1.Cells(lnLineaHoja1, 9).Formula - xlsHoja1.Cells(lnLineaHoja1 - 1, 9).Formula
    'VAR ACUMULADA HOJA001
    xlsHoja1.Cells(lnLineaHoja1 + 3, 9).Formula = xlsHoja1.Cells(lnLineaHoja1, 9).Formula - lnSaldoEval
    xlsHoja1.Cells(lnLineaHoja1 + 3, 10).Formula = xlsHoja1.Cells(lnLineaHoja1, 9).Formula / lnSaldoEval - 1
    '*************HOJA001
    'PIE DE PAGINA HOJA001
    lnLineaHoja1 = lnLineaHoja1 + 1
    xlsHoja1.Range("B" & Trim(Str(lnLineaHoja1)) & ":" & "J" & Trim(Str(lnLineaHoja1))).Borders.LineStyle = 1
    xlsHoja1.Range("B" & CStr(lnLineaHoja1)) = "PARTICIPACION %"
    xlsHoja1.Range("B" & CStr(lnLineaHoja1), "B" & CStr(lnLineaHoja1 + 2)).Interior.Color = RGB(199, 199, 199)
    xlsHoja1.Range("C" & CStr(lnLineaHoja1), "D" & CStr(lnLineaHoja1)).MergeCells = True
    xlsHoja1.Range("E" & CStr(lnLineaHoja1), "F" & CStr(lnLineaHoja1)).MergeCells = True
    xlsHoja1.Range("G" & CStr(lnLineaHoja1), "H" & CStr(lnLineaHoja1)).MergeCells = True
    xlsHoja1.Range("I" & CStr(lnLineaHoja1), "J" & CStr(lnLineaHoja1)).MergeCells = True
    xlsHoja1.Cells(lnLineaHoja1, 3).Formula = xlsHoja1.Cells(lnLineaHoja1 - 1, 3).Formula / xlsHoja1.Cells(lnLineaHoja1 - 1, 9).Formula
    xlsHoja1.Cells(lnLineaHoja1, 5).Formula = xlsHoja1.Cells(lnLineaHoja1 - 1, 5).Formula / xlsHoja1.Cells(lnLineaHoja1 - 1, 9).Formula
    xlsHoja1.Cells(lnLineaHoja1, 7).Formula = xlsHoja1.Cells(lnLineaHoja1 - 1, 7).Formula / xlsHoja1.Cells(lnLineaHoja1 - 1, 9).Formula
    xlsHoja1.Range("I" & lnLineaHoja1) = "= C" & lnLineaHoja1 & "+E" & lnLineaHoja1 & "+G" & lnLineaHoja1
    'xlsHoja1.Cells(lnLineaHoja1, 9).Formula = xlsHoja1.Cells(lnLineaHoja1, 3).Formula + xlsHoja1.Cells(lnLineaHoja1, 5).Formula + xlsHoja1.Cells(lnLineaHoja1, 7).Formula
    xlsHoja1.Range("C" & lnLineaHoja1 & ":J" & lnLineaHoja1).NumberFormat = "0.00%"
    
    lnLineaHoja1 = lnLineaHoja1 + 1
    xlsHoja1.Range("B" & Trim(Str(lnLineaHoja1)) & ":" & "J" & Trim(Str(lnLineaHoja1))).Borders.LineStyle = 1
    xlsHoja1.Range("B" & CStr(lnLineaHoja1)) = "VAR. MENSUAL"
    xlsHoja1.Range("C" & CStr(lnLineaHoja1), "D" & CStr(lnLineaHoja1)).MergeCells = True
    xlsHoja1.Range("E" & CStr(lnLineaHoja1), "F" & CStr(lnLineaHoja1)).MergeCells = True
    xlsHoja1.Range("G" & CStr(lnLineaHoja1), "H" & CStr(lnLineaHoja1)).MergeCells = True
    xlsHoja1.Range("I" & CStr(lnLineaHoja1), "J" & CStr(lnLineaHoja1)).MergeCells = True
    xlsHoja1.Range("C" & CStr(lnLineaHoja1) & ":J" & CStr(lnLineaHoja1)).NumberFormat = "#,##0"
    
    lnLineaHoja1 = lnLineaHoja1 + 1
    xlsHoja1.Range("B" & Trim(Str(lnLineaHoja1)) & ":" & "J" & Trim(Str(lnLineaHoja1))).Borders.LineStyle = 1
    xlsHoja1.Range("B" & CStr(lnLineaHoja1)) = "VAR. ACUMULADA"
    
    xlsHoja1.Range("C" & CStr(lnLineaHoja1) & ":C" & CStr(lnLineaHoja1)).NumberFormat = "#,##0"
    xlsHoja1.Range("D" & CStr(lnLineaHoja1) & ":D" & CStr(lnLineaHoja1)).NumberFormat = "0.00%"
    xlsHoja1.Range("E" & CStr(lnLineaHoja1) & ":E" & CStr(lnLineaHoja1)).NumberFormat = "#,##0"
    xlsHoja1.Range("F" & CStr(lnLineaHoja1) & ":F" & CStr(lnLineaHoja1)).NumberFormat = "0.00%"
    xlsHoja1.Range("G" & CStr(lnLineaHoja1) & ":G" & CStr(lnLineaHoja1)).NumberFormat = "#,##0"
    xlsHoja1.Range("H" & CStr(lnLineaHoja1) & ":H" & CStr(lnLineaHoja1)).NumberFormat = "0.00%"
    xlsHoja1.Range("I" & CStr(lnLineaHoja1) & ":I" & CStr(lnLineaHoja1)).NumberFormat = "#,##0"
    xlsHoja1.Range("J" & CStr(lnLineaHoja1) & ":J" & CStr(lnLineaHoja1)).NumberFormat = "0.00%"
    
    Set rsAhorros = Nothing
    Set oAgencia = Nothing
    Set oCaja = Nothing
End Sub
Private Sub generaHojaExcelAgencia_RptDepositosTotales(ByRef xlsHoja2 As Worksheet, ByVal psAgeCod As String, ByVal psNomHoja2 As String, ByVal psAnio As String, ByVal psMes As String, ByVal psAnioHasta As String, ByVal psMesHasta As String)
    Dim oBalance As New DbalanceCont
    Dim oCaja As New DCajaGeneral
    Dim rsAhorros As New ADODB.Recordset
    Dim oAgencia As New DAgencia
    Dim cAgencia As String
    Dim lnLineaActual As Integer
    Dim ldFecha As Date
    Dim lnSaldoEval As Currency
    Dim lnVarEval As Double
    Dim lcFila As Integer
    Dim i As Integer
    Dim lnFilaInicio As Integer
    i = 1
    
    ldFecha = CDate("31/12/" & Format(Val(txtAnioHasta.Text) - 1, "0000"))
    
    cAgencia = UCase(Trim(oAgencia.GetAgencias(psAgeCod)))
    
    xlsHoja2.Cells.Font.Name = "Arial"
    xlsHoja2.Cells.Font.Size = 9
    xlsHoja2.Range("A:A").ColumnWidth = 4
    xlsHoja2.Range("B:B").ColumnWidth = 10
    xlsHoja2.Range("C:C").ColumnWidth = 12
    xlsHoja2.Range("D:D").ColumnWidth = 16
    xlsHoja2.Range("E:E").ColumnWidth = 16
    xlsHoja2.Range("F:F").ColumnWidth = 16
    xlsHoja2.Range("G:G").ColumnWidth = 16
    xlsHoja2.Range("H:H").ColumnWidth = 16
    xlsHoja2.Range("I:I").ColumnWidth = 16
    
    xlsHoja2.Range("B3", "I3").MergeCells = True
    xlsHoja2.Range("B3", "I3").HorizontalAlignment = xlCenter
    xlsHoja2.Range("B3", "I6").Font.Bold = True
    xlsHoja2.Range("B3:I3").Font.Size = 14
    xlsHoja2.Range("B3:I3").Interior.Color = RGB(255, 255, 0)
    xlsHoja2.Range("B3") = "DEPOSITOS TOTALES"
    xlsHoja2.Range("B4", "C4").MergeCells = True
    xlsHoja2.Range("B4") = cAgencia
    xlsHoja2.Range("I4", "I4").HorizontalAlignment = xlRight
    xlsHoja2.Range("I4") = "Año " & psAnioHasta
    
    xlsHoja2.Range("B5", "C5").MergeCells = True
    xlsHoja2.Range("B5", "C5").HorizontalAlignment = xlLeft
    xlsHoja2.Range("B5") = " ( en nuevo soles ) "
    
    xlsHoja2.Range("G5", "I5").MergeCells = True
    xlsHoja2.Range("G5", "I5").HorizontalAlignment = xlRight
    xlsHoja2.Range("G5") = "RESULTADOS VS PROYECCIONES"
    
    'CABECERA HOJA002
    xlsHoja2.Range("B6", "I6").HorizontalAlignment = xlCenter
    xlsHoja2.Range("B6:I6").Interior.Color = RGB(255, 255, 0)
    xlsHoja2.Range("B6") = "Meses"
    xlsHoja2.Range("C6") = "# Ctas."
    xlsHoja2.Range("D6") = "Importe S/."
    xlsHoja2.Range("E6") = "Crecimiento anual / mensual"
    xlsHoja2.Range("F6") = "Variaciones %"
    xlsHoja2.Range("G6") = "Proyectado"
    xlsHoja2.Range("H6") = "Resultados"
    xlsHoja2.Range("I6") = "Avance %"

    xlsHoja2.Range("B" & Trim(Str(6)) & ":" & "I" & Trim(Str(6))).Borders.LineStyle = 1
    lnLineaActual = 6
    lnFilaInicio = lnLineaActual + 1
    '*********** DEPOSITOS TOTALES
    Set rsAhorros = oCaja.ReporteDepositosTotales(psAnio, psMes, psAnioHasta, psMesHasta, psAgeCod)
    Do While Not rsAhorros.EOF
        If i = 1 Then
            rsAhorros.MoveNext
            i = i + 1
        Else
            xlsHoja2.Range("B" & Trim(Str(lnLineaActual)) & ":" & "I" & Trim(Str(lnLineaActual))).Borders.LineStyle = 1
            lnLineaActual = lnLineaActual + 1
            xlsHoja2.Cells(lnLineaActual, 2).Formula = Format(rsAhorros!cPeriodo, "dd/mm/yyyy")
            xlsHoja2.Cells(lnLineaActual, 3).Formula = rsAhorros!nNCuentas
            xlsHoja2.Cells(lnLineaActual, 4).Formula = rsAhorros!mImporteTotal
            xlsHoja2.Cells(lnLineaActual, 5).Formula = rsAhorros!mCrecimiento
            xlsHoja2.Cells(lnLineaActual, 6).Formula = rsAhorros!fVarMensual / 100
            xlsHoja2.Cells(lnLineaActual, 7).Formula = rsAhorros!Proyectado
            xlsHoja2.Cells(lnLineaActual, 8).Formula = rsAhorros!Resultado
            xlsHoja2.Cells(lnLineaActual, 9).Formula = rsAhorros!Avance / 100
            If CDate(rsAhorros!cPeriodo) = ldFecha Then
                lnVarEval = rsAhorros!mImporteTotal
                lcFila = lnLineaActual + 1
            End If
            If CDate(rsAhorros!cPeriodo) > ldFecha Then
                lnSaldoEval = lnSaldoEval + rsAhorros!mCrecimiento
            End If
            rsAhorros.MoveNext
        End If
    Loop
    xlsHoja2.Range("D" & CStr(lnFilaInicio) & ":E" & CStr(lnLineaActual)).NumberFormat = "#,##0"
    xlsHoja2.Range("G" & CStr(lnFilaInicio) & ":H" & CStr(lnLineaActual)).NumberFormat = "#,##0"
    xlsHoja2.Range("F" & CStr(lnFilaInicio) & ":F" & CStr(lnLineaActual)).NumberFormat = "0.00%"
    xlsHoja2.Range("I" & CStr(lnFilaInicio) & ":I" & CStr(lnLineaActual)).NumberFormat = "0.00%"
    
    xlsHoja2.Range("B" & Trim(Str(lnLineaActual)) & ":" & "I" & Trim(Str(lnLineaActual))).Borders.LineStyle = 1
    lnLineaActual = lnLineaActual + 1
    xlsHoja2.Range("C" & CStr(lnLineaActual), "D" & CStr(lnLineaActual)).MergeCells = True
    xlsHoja2.Range("C" & CStr(lnLineaActual)) = "VARIACION ACUMULADA"
    'VARIACION ACUMULADA
    xlsHoja2.Range("C" & CStr(lnLineaActual), "C" & CStr(lnLineaActual)).Interior.Color = RGB(199, 199, 199)
    xlsHoja2.Cells(lnLineaActual, 5).Formula = lnSaldoEval
    xlsHoja2.Cells(lnLineaActual, 6).Formula = (xlsHoja2.Cells(lnLineaActual - 1, 4).Formula / lnVarEval - 1) * 100
    xlsHoja2.Range("B" & Trim(Str(lnLineaActual)) & ":" & "I" & Trim(Str(lnLineaActual))).Borders.LineStyle = 1
    'FORMULA
    xlsHoja2.Range("F" & lnLineaActual) = "= D" & lnLineaActual - 1 & "/D" & lcFila - 1 & "-1"
    xlsHoja2.Range("F" & lnLineaActual & ":F" & lnLineaActual).NumberFormat = "0.00%"
    'Crear Grafico
    Set oChart = xlsHoja2.ChartObjects.Add(50, 13 * lnLineaActual, lnLineaActual * 25, 300).Chart
    oChart.SetSourceData Source:=xlsHoja2.Range("'" & psNomHoja2 & "'!$B$7:$B$" & lnLineaActual - 1 & ",'" & psNomHoja2 & "'!$G$7:$G$" & lnLineaActual - 1 & ",'" & psNomHoja2 & "'!$H$7:$H$" & lnLineaActual - 1 & "")
    
    Set rsAhorros = Nothing
    Set oAgencia = Nothing
    Set oCaja = Nothing
End Sub



Private Sub OptAnalista_Click(Index As Integer)
Dim bCheck As Boolean
Dim i As Integer
    If Index = 0 Then
        bCheck = True
    Else
        bCheck = False
    End If
    If lstAge.ListCount <= 0 Then
        Exit Sub
    End If
    For i = 0 To lstAge.ListCount - 1
        lstAge.Selected(i) = bCheck
    Next i
End Sub

Private Sub txtAnioHasta_Change()
If Len(Me.txtAnioHasta.Text) = 4 Then
        Me.cboMesHasta.SetFocus
    End If
End Sub
Private Sub txtAnioHasta_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub
Private Sub txtAnio_Change()
    If Len(txtAnio.Text) = 4 Then
        cboMes.SetFocus
    End If
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Public Sub Ini(ByVal psOpeCod As String, ByVal psOpeDesc As String)
    'fsOpeCod = psOpeCod
    'Me.Caption = psOpeDesc
    Me.Show 1
End Sub
Private Sub Form_Load()
CentraForm Me
Call CargarAgencias
End Sub
Private Sub CargarAgencias()
    Dim sqlAge As String
    Dim rsAge As ADODB.Recordset
    Dim oCon As DConecta

    Set oCon = New DConecta

    oCon.AbreConexion
    Set rsAge = New ADODB.Recordset
    sqlAge = "Select cAgeDescripcion cNomtab, cAgeCod cValor From Agencias where nEstado = 1"
    Set rsAge = oCon.CargaRecordSet(sqlAge)

    lstAge.Clear

    If Not RSVacio(rsAge) Then
        While Not rsAge.EOF
            lstAge.AddItem Trim(rsAge!cNomtab) & Space(500) & Trim(rsAge!cValor)
            rsAge.MoveNext
        Wend
        'lstAge.AddItem "Todos" & Space(500) & "TODOS"
    End If

    rsAge.Close
    Set rsAge = Nothing
End Sub


