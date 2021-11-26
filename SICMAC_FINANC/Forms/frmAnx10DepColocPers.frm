VERSION 5.00
Begin VB.Form frmAnx10DepColocPers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anexo 10: Depositos, Colocaciones y Personal por Oficinas"
   ClientHeight    =   4440
   ClientLeft      =   1350
   ClientTop       =   2340
   ClientWidth     =   7350
   Icon            =   "frmAnx10DepColocPers.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   405
      Left            =   2955
      TabIndex        =   4
      Top             =   5895
      Width           =   1185
   End
   Begin VB.Frame fraPeriodo 
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
      Height          =   750
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   4965
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3660
         MaxLength       =   4
         TabIndex        =   2
         Top             =   270
         Width           =   1095
      End
      Begin VB.ComboBox cboMes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmAnx10DepColocPers.frx":030A
         Left            =   870
         List            =   "frmAnx10DepColocPers.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   3090
         TabIndex        =   11
         Top             =   330
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   285
         TabIndex        =   10
         Top             =   330
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4770
      TabIndex        =   6
      Top             =   3945
      Width           =   1185
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   6000
      TabIndex        =   7
      Top             =   3945
      Width           =   1185
   End
   Begin Sicmact.FlexEdit fg 
      Height          =   2985
      Left            =   120
      TabIndex        =   3
      Top             =   870
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5265
      Cols0           =   10
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Agencia-Pais-Depa-Prov-Distr-Tipo Oficina-Codigo-cAgeCod-cTipo"
      EncabezadosAnchos=   "350-2200-500-550-500-500-1400-700-1-1"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-2-3-4-5-6-7-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-3-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C-C-C-C-L-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   405
      Left            =   150
      TabIndex        =   5
      Top             =   3945
      Width           =   1185
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   405
      Left            =   150
      TabIndex        =   8
      Top             =   3945
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   1425
      TabIndex        =   9
      Top             =   3945
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "frmAnx10DepColocPers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsArchivo As String
Dim lbExcel As Boolean
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim ldFecha  As Date
Dim oBarra As clsProgressBar
Dim oCon As DConecta
Dim sSql As String
Dim rs   As ADODB.Recordset
Dim nNroLocales As Integer
Dim opeGraba As Integer

Private Sub TotalizaColumnas(pnColTot As Integer, pnCol1 As Integer, pnCol2 As Integer, Optional pnCol3 As Integer = 0, Optional pnCol4 As Integer = 0)
Dim K As Integer
Dim lsFormula As String
For K = 1 To (fg.Rows - 1) Step 1
    lsFormula = ""
    If pnCol3 > 0 Then
        lsFormula = lsFormula & "+" & xlHoja1.Range(xlHoja1.Cells(K + 10, pnCol3), xlHoja1.Cells(K + 10, pnCol3)).Address
    End If
    If pnCol4 > 0 Then
        lsFormula = lsFormula & "+" & xlHoja1.Range(xlHoja1.Cells(K + 10, pnCol4), xlHoja1.Cells(K + 10, pnCol4)).Address
    End If
    xlHoja1.Range(xlHoja1.Cells(K + 10, pnColTot), xlHoja1.Cells(K + 10, pnColTot)).Formula = "=+" & xlHoja1.Range(xlHoja1.Cells(K + 10, pnCol1), xlHoja1.Cells(K + 10, pnCol1)).Address & "+" & xlHoja1.Range(xlHoja1.Cells(K + 10, pnCol2), xlHoja1.Cells(K + 10, pnCol2)).Address & lsFormula
Next
End Sub

Private Sub CabeceraAnexo10(pdFecha As Date, Optional lbSoles As Boolean = True)
   Dim lbExisteHoja  As Boolean
   Dim I  As Long
   Dim lnFila As Integer
   On Error Resume Next
   
     
   'TITULOS GENERALES
   '=================
    
    xlHoja1.Cells(1, 2) = "SUPERINTENDENCIA DE BANCA Y SEGUROS"
    xlHoja1.Cells(1, 22) = "ANEXO Nº 10"
    xlHoja1.Cells(3, 8) = ImpreFormat("EMPRESA : " & UCase(gsNomCmac), 100) & "CODIGO : " & gsCodCMAC
    xlHoja1.Cells(4, 8) = "DEPOSITOS, COLOCACIONES Y PERSONAL POR OFICINAS"
    xlHoja1.Cells(5, 1) = "Fecha : Al " & Format(pdFecha, "dd mmmm yyyy")
     
    If lbSoles = False Then
        '''xlHoja1.Cells(6, 1) = "( En Nuevos Soles )" 'MARG ERS044-2016
        xlHoja1.Cells(6, 1) = "( En " & StrConv(gcPEN_PLURAL, vbProperCase) & " )" 'MARG ERS044-2016
    Else
        '''xlHoja1.Cells(6, 1) = "( En Miles de Nuevos Soles )" 'MARG ERS044-2016
        xlHoja1.Cells(6, 1) = "( En Miles de " & StrConv(gcPEN_PLURAL, vbProperCase) & " )" 'MARG ERS044-2016
    End If
         
    xlHoja1.Range("B1:H1").MergeCells = True
    xlHoja1.Range("H3:P3").MergeCells = True
    xlHoja1.Range("H4:P4").MergeCells = True
    xlHoja1.Range("A5:V5").MergeCells = True
    xlHoja1.Range("A6:V6").MergeCells = True
        
    xlHoja1.Range("A4:V4").HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A5:V5").HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A6:V6").HorizontalAlignment = xlHAlignCenter
    
    'TITULOS DE LOS ENCABEZADOS
    '==========================
    xlHoja1.Cells(8, 1) = "Oficina"
    xlHoja1.Cells(8, 2) = "Código Ubicación Geográfica"
    xlHoja1.Cells(8, 6) = "Oficina"
    xlHoja1.Cells(8, 8) = "Depósitos"
    xlHoja1.Cells(8, 15) = "Colocaciones"
    xlHoja1.Cells(8, 18) = "Personal 2/"
    xlHoja1.Cells(8, 23) = "Nro. de Cajeros Automáticos"
    
    xlHoja1.Cells(9, 8) = "Ahorros"
    xlHoja1.Cells(9, 11) = "Plazo"
    xlHoja1.Cells(9, 14) = "Total"
    xlHoja1.Cells(9, 15) = "Moneda Nacional"
    xlHoja1.Cells(9, 16) = "Equivalente ME"
    xlHoja1.Cells(9, 17) = "Total"
    xlHoja1.Cells(9, 18) = "Gerentes"
    xlHoja1.Cells(9, 19) = "Funcionarios"
    xlHoja1.Cells(9, 20) = "Empleados"
    xlHoja1.Cells(9, 21) = "Otros"
    xlHoja1.Cells(9, 22) = "Total"
     
    xlHoja1.Cells(10, 2) = "Pais"
    xlHoja1.Cells(10, 3) = "Dpto"
    xlHoja1.Cells(10, 4) = "Prov"
    xlHoja1.Cells(10, 5) = "Dist"
    xlHoja1.Cells(10, 6) = "Tipo 1/"
    xlHoja1.Cells(10, 7) = "Codigo"
    xlHoja1.Cells(10, 8) = "Moneda Nacional"
    xlHoja1.Cells(10, 9) = "Equivalente ME"
    xlHoja1.Cells(10, 10) = "Total"
    xlHoja1.Cells(10, 11) = "Moneda Nacional"
    xlHoja1.Cells(10, 12) = "Equivalente ME"
    xlHoja1.Cells(10, 13) = "Total"
    
    xlHoja1.Range("A8:A10").MergeCells = True
    xlHoja1.Range("B8:E9").MergeCells = True
    xlHoja1.Range("F8:G8").MergeCells = True
    xlHoja1.Range("F9:F10").MergeCells = True
    xlHoja1.Range("G9:G10").MergeCells = True
    xlHoja1.Range("H8:N8").MergeCells = True
    xlHoja1.Range("H9:J9").MergeCells = True
    xlHoja1.Range("K9:M9").MergeCells = True
    xlHoja1.Range("N9:N10").MergeCells = True
    xlHoja1.Range("O8:Q8").MergeCells = True
    xlHoja1.Range("O9:O10").MergeCells = True
    xlHoja1.Range("P9:P10").MergeCells = True
    xlHoja1.Range("Q9:Q10").MergeCells = True
    xlHoja1.Range("R8:V8").MergeCells = True
    xlHoja1.Range("R9:R10").MergeCells = True
    xlHoja1.Range("S9:S10").MergeCells = True
    xlHoja1.Range("T9:T10").MergeCells = True
    xlHoja1.Range("U9:U10").MergeCells = True
    xlHoja1.Range("V9:V10").MergeCells = True
    xlHoja1.Range("W8:W10").MergeCells = True
    
    'Poner negrita a las cabeceras
    xlHoja1.Range("A1:W6").Font.Bold = True
    xlHoja1.Range("A8:W10").Font.Bold = True
    
    xlHoja1.Cells.Select
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 8
    xlHoja1.Cells.EntireColumn.AutoFit
    'xlHoja1.Range("H4,H4").Font.Size = 12
    xlHoja1.Range("A3:V3").Font.Size = 10
    If gsCodCMAC = "102" Then
        xlHoja1.Range("A1:A1").ColumnWidth = 0
    Else
        xlHoja1.Range("A1:A1").ColumnWidth = 16
    End If
    xlHoja1.Range("W1:W1").ColumnWidth = 10
    xlHoja1.Range("N1:N1").ColumnWidth = 11
    xlHoja1.Range("J1:J1").ColumnWidth = 11
    xlHoja1.Range("Q1:Q1").ColumnWidth = 11
     
    xlHoja1.Range("A8:W10").WrapText = True
    xlHoja1.Range("A8:W10").HorizontalAlignment = xlCenter
    xlHoja1.Range("A8:W10").VerticalAlignment = xlCenter
        
    'Ponerle marco a los encabezados
    With xlHoja1.Range("A8:W10")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
     
    ''Lleno las cabeceras de las filas : Los nombres de las agencias
    lnFila = 10
    
    Set rs = CargaDatosAnexo("Titulos", 0)
    Do While Not rs.EOF
        lnFila = lnFila + 1
        If Not gsCodCMAC = "102" Then
            xlHoja1.Cells(lnFila, 1) = rs!cAgeDescripcion
        End If
        xlHoja1.Cells(lnFila, 2) = rs!cPais
        xlHoja1.Cells(lnFila, 3) = rs!cDepa
        xlHoja1.Cells(lnFila, 4) = rs!cProv
        xlHoja1.Cells(lnFila, 5) = rs!cDist
        xlHoja1.Cells(lnFila, 6) = rs!cTipo
        xlHoja1.Cells(lnFila, 7) = rs!cCodigo
        rs.MoveNext
    Loop
    
    'Cuadricula principal
    ExcelCuadro xlHoja1, 1, 11, 23, 10 + fg.Rows - 1
     
    'Total
    xlHoja1.Cells(10 + fg.Rows, 1) = "Total:"
    xlHoja1.Range("A" & Trim(Str(10 + fg.Rows)) & ":G" & Trim(Str(10 + fg.Rows)) & "").MergeCells = True
    
    'Formulas de Total
    For I = 8 To 23
        xlHoja1.Cells(10 + fg.Rows, I).Formula = "=SUM(" & Trim(Chr(96 + I)) & "11:" & Trim(Chr(96 + I)) & "" & Trim(Str(10 + fg.Rows - 1)) & ")"
        xlHoja1.Range(ExcelColumnaString(CInt(I)) & (10 + fg.Rows) & ":" & ExcelColumnaString(CInt(I)) & (10 + fg.Rows)).NumberFormat = "#,##0.00"
    Next
     
    xlHoja1.Range("A" & Trim(Str(10 + fg.Rows)) & ":W" & Trim(Str(10 + fg.Rows)) & "").Font.Bold = True
    
    'Bordes de Total
    With xlHoja1.Range("A" & Trim(Str(10 + fg.Rows)) & ":W" & Trim(Str(10 + fg.Rows)) & "")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    lnFila = 10 + fg.Rows + 1
    
    'Glosa
    xlHoja1.Cells(lnFila, 1) = "1/"
    xlHoja1.Cells(lnFila, 2) = "Considere(1) Oficina Principal; (2) Sucursal; (3) Agencia; (4) Oficina Especial; (5) Local Compartido. En caso de operar en un local compartido con otra entidad del sistema financiero"
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 2) = "o de seguros, marcar un asterisco(*) al costado del número indicativo del tipo de oficina"
    lnFila = lnFila + 1
    xlHoja1.Cells(lnFila, 1) = "2/"
    xlHoja1.Cells(lnFila, 2) = "Debe considerarse a todo el personal, de acuerdo con la función que desempeñan, independientemente de si son nombrados, contratos o contratados por terceros"
    
    xlHoja1.Range("A" & Trim(Str(lnFila - 2)) & ":A" & Trim(Str(lnFila)) & "").HorizontalAlignment = xlRight
     
    'Tamaño de letra de la firma
    xlHoja1.Range("A" & Trim(Str(lnFila - 2)) & ":W" & Trim(Str(lnFila)) & "").Font.Size = 7
    xlHoja1.Range("A" & Trim(Str(lnFila - 2)) & ":W" & Trim(Str(lnFila)) & "").Font.Bold = True
    
    'Firmas
    '======
    lnFila = lnFila + 6
    xlHoja1.Cells(lnFila, 2) = "GERENTE GENERAL"
    xlHoja1.Range("B" & Trim(Str(lnFila)) & ":F" & Trim(Str(lnFila)) & "").MergeCells = True
    
    xlHoja1.Cells(lnFila, 12) = "CONTADOR GENERAL"
    xlHoja1.Range("L" & Trim(Str(lnFila)) & ":M" & Trim(Str(lnFila)) & "").MergeCells = True
    
    xlHoja1.Cells(lnFila, 20) = "HECHO POR"
    xlHoja1.Range("S" & Trim(Str(lnFila)) & ":T" & Trim(Str(lnFila)) & "").MergeCells = True
    
    lnFila = lnFila + 1
    
    xlHoja1.Cells(lnFila, 12) = "MATRICULA NRO"
    xlHoja1.Range("L" & Trim(Str(lnFila)) & ":M" & Trim(Str(lnFila)) & "").MergeCells = True
    
    xlHoja1.Range("A" & Trim(Str(lnFila - 1)) & ":W" & Trim(Str(lnFila)) & "").Font.Bold = True
    xlHoja1.Range("A" & Trim(Str(lnFila - 1)) & ":W" & Trim(Str(lnFila)) & "").HorizontalAlignment = xlCenter
    
    'Rayas de las firmas
    xlHoja1.Range("B" & Trim(Str(lnFila - 1)) & ":F" & Trim(Str(lnFila - 1)) & "").Borders(xlEdgeTop).LineStyle = xlContinuous
     xlHoja1.Range("B" & Trim(Str(lnFila - 1)) & ":F" & Trim(Str(lnFila - 1)) & "").Borders(xlEdgeTop).Weight = xlThin
     xlHoja1.Range("L" & Trim(Str(lnFila - 1)) & ":M" & Trim(Str(lnFila - 1)) & "").Borders(xlEdgeTop).LineStyle = xlContinuous
     xlHoja1.Range("L" & Trim(Str(lnFila - 1)) & ":M" & Trim(Str(lnFila - 1)) & "").Borders(xlEdgeTop).Weight = xlThin
     xlHoja1.Range("S" & Trim(Str(lnFila - 1)) & ":T" & Trim(Str(lnFila - 1)) & "").Borders(xlEdgeTop).LineStyle = xlContinuous
     xlHoja1.Range("S" & Trim(Str(lnFila - 1)) & ":T" & Trim(Str(lnFila - 1)) & "").Borders(xlEdgeTop).Weight = xlThin
 Exit Sub
errCabecera:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub ActivaBotonesEditar(pbActiva As Boolean)
cmdGrabar.Visible = pbActiva
cmdCancelar.Visible = pbActiva
cmdEditar.Visible = Not pbActiva
cmdNuevo.Visible = Not pbActiva
cmdGenerar.Enabled = Not pbActiva
fg.lbEditarFlex = pbActiva
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtAnio.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
ActivaBotonesEditar False
CargaDatosUbigeo
fg.SetFocus
End Sub

Private Sub cmdEditar_Click()
opeGraba = 2
ActivaBotonesEditar True
fg.SetFocus
End Sub

Private Sub cmdGenerar_Click()
On Error GoTo GeneraEstadError
Dim nTipoCambio As Currency
Dim I As Integer
Dim nMiles As Boolean
Dim oTC As New nTipoCambio
    
   ldFecha = CDate("01/" & Format(Me.cboMes.ListIndex + 1, "00") & "/" & Format(txtAnio, "0000"))
   ldFecha = DateAdd("m", 1, ldFecha) - 1
    
   nTipoCambio = oTC.EmiteTipoCambio(ldFecha + 1, TCFijoMes)
   Set oTC = Nothing
    
   'lsArchivo = App.path & "\SPOOLER\" & "Anx10_" & Format(ldFecha & " " & Time, "mmddyyyyhhmmss") & ".XLS" Comentado PASIERS1332014
   lsArchivo = App.path & "\SPOOLER\" & "Anx10_" & Format(ldFecha & " " & Time, "mmddyyyyhhmmss") & ".xlsx" 'PASIERS1332014
   
   lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If lbExcel Then
    
        Set oBarra = New clsProgressBar
        oBarra.ShowForm frmReportes
        oBarra.Max = 4
        oBarra.Progress 0, "ANEXO 10: DEPOSITOS, COLOCACIONES Y PERSONAL POR OFICINAS", "Generando Reporte", "", vbBlue
    
        For I = 2 To 1 Step -1
            If I = 1 Then
                '''ExcelAddHoja "NuevosSoles", xlLibro, xlHoja1 'MARG ERS044-2016
                ExcelAddHoja StrConv(gcPEN_PLURAL, vbProperCase), xlLibro, xlHoja1 'MARG ERS044-2016
                nMiles = False
            ElseIf I = 2 Then
                '''ExcelAddHoja "MilesNuevosSoles", xlLibro, xlHoja1 'MARG ERS044-2016
                ExcelAddHoja "Miles" & StrConv(gcPEN_PLURAL, vbProperCase), xlLibro, xlHoja1 'MARG ERS044-2016
                nMiles = True
            End If
            CabeceraAnexo10 ldFecha, nMiles
            MuestraDatosAnexo "Ahorros", 8, 1, nTipoCambio, nMiles, "#,###,###.00"
            MuestraDatosAnexo "Ahorros", 9, 2, nTipoCambio, nMiles, "#,###,###.00"
            TotalizaColumnas 10, 8, 9
            MuestraDatosAnexo "Plazo", 11, 1, nTipoCambio, nMiles, "#,###,###.00"
            MuestraDatosAnexo "Plazo", 12, 2, nTipoCambio, nMiles, "#,###,###.00"
            TotalizaColumnas 13, 11, 12
            TotalizaColumnas 14, 10, 13
            MuestraDatosAnexo "Creditos", 15, 1, nTipoCambio, nMiles, "#,###,###.00"
            MuestraDatosAnexo "Creditos", 16, 2, nTipoCambio, nMiles, "#,###,###.00"
            TotalizaColumnas 17, 15, 16
            MuestraDatosAnexo "Gerentes", 18, 1, nTipoCambio, False, "#,###,###"
            MuestraDatosAnexo "Funcionarios", 19, 1, nTipoCambio, False, "#,###,###"
            MuestraDatosAnexo "Empleados", 20, 1, nTipoCambio, False, "#,###,###"
            MuestraDatosAnexo "Otros", 21, 1, nTipoCambio, False, "#,###,###"
            TotalizaColumnas 22, 18, 19, 20, 21
            
'            If gsCodCMAC = "102" Then
                xlHoja1.Cells(8, 8) = ""
                xlHoja1.Range("H1:J1").EntireColumn.Insert False
                xlHoja1.Range("H8:N8").Merge True
                xlHoja1.Range("H8:N8").BorderAround
                xlHoja1.Range("H9:J9").Merge True
                xlHoja1.Range("H8:J10").Font.Bold = True
                xlHoja1.Cells(9, 8) = "VISTA"
                xlHoja1.Cells(10, 8) = "M.Nac."
                xlHoja1.Cells(10, 9) = "Eq ME"
                xlHoja1.Cells(10, 10) = "Total"
                xlHoja1.Range("H9:J10").HorizontalAlignment = xlCenter
                xlHoja1.Cells(8, 8) = "Depósitos"
'            End If
        Next

        oBarra.Progress 100, "ANEXO 10: DEPOSITOS, COLOCACIONES Y PERSONAL POR OFICINAS", "REPORTE TERMINADO", "", vbBlue
        oBarra.CloseForm frmReportes
        Set oBarra = Nothing
        
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
    End If
    
    If lsArchivo <> "" Then
       CargaArchivo lsArchivo, App.path & "\SPOOLER\"
    End If
    
Exit Sub
GeneraEstadError:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    If Not oBarra Is Nothing Then
        oBarra.CloseForm frmReportes
        Set oBarra = Nothing
    End If
    If lbExcel = True Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Sub

Private Sub MuestraDatosAnexo(psTipo As String, pnCol As Integer, pnMoneda As Integer, pnTipoCambio As Currency, pnMiles As Boolean, pnFormato As String)
Dim lnFila As Integer
lnFila = 11
Set rs = CargaDatosAnexo(psTipo, pnMoneda)

Do While Not rs.EOF
    If pnMoneda = gMonedaNacional Then
        If pnMiles = True Then
            xlHoja1.Cells(lnFila, pnCol) = Format(rs!nSaldo / 1000, pnFormato)
        Else
            xlHoja1.Cells(lnFila, pnCol) = Format(rs!nSaldo, pnFormato)
        End If
    ElseIf pnMoneda = gMonedaExtranjera Then
        If psTipo = "Creditos" Then 'Agregado no se multiplica por t.c.
            If pnMiles = True Then
                xlHoja1.Cells(lnFila, pnCol) = Format(rs!nSaldo / 1000, pnFormato)
            Else
                xlHoja1.Cells(lnFila, pnCol) = Format(rs!nSaldo, pnFormato)
            End If
        Else
            If pnMiles = True Then
                xlHoja1.Cells(lnFila, pnCol) = Format(rs!nSaldo / 1000 * pnTipoCambio, pnFormato)
            Else
                xlHoja1.Cells(lnFila, pnCol) = Format(rs!nSaldo * pnTipoCambio, pnFormato)
            End If
        End If
    End If
    
    lnFila = lnFila + 1
    rs.MoveNext
Loop
End Sub

Private Function CargaDatosAnexo(psTipo As String, pnMoneda As Integer) As ADODB.Recordset

Dim lbAhorrosRestringidos As Boolean
lbAhorrosRestringidos = False
 
Dim sservidorconsolidada As String
Dim rCargaRuta As ADODB.Recordset

Dim lsWhere1 As String
Dim lsWhere2 As String

Dim lsWhere1C As String
Dim lsWhere2C As String
 
    Set rCargaRuta = oCon.CargaRecordSet("select nconssisvalor from constsistema where nconssiscod=" & gConstSistServCentralRiesgos)
    If rCargaRuta.BOF Then
    Else
        sservidorconsolidada = rCargaRuta!nConsSisValor
    End If
    Set rCargaRuta = Nothing
'JOEP20190531 Mejora Anexo10

    sSql = "Exec stp_sel_Anexo10ObtieneDatos '" & Format(ldFecha, "yyyy/mm/dd") & "','" & psTipo & "','" & pnMoneda & "'," & IIf(gbBitCentral, 1, 0) & "," & IIf(lbAhorrosRestringidos, 1, 0) & ",'" & gCapAhorros & "','" & gCapPlazoFijo & "','" & gCapCTS & "','" & gsCodCMAC & "'"

'    Select Case psTipo
'
'    Case "Titulos"
'        'Centralizado y Distribuido
'        sSql = "select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, A10.cTipo, A10.cCodigo "
'        sSql = sSql & " from agencias A "
'        sSql = sSql & " Left join anx10datos A10 on A.cAgeCod=A10.cAgeCod "
'        sSql = sSql & " Where bcambiarAge=0 "
'        sSql = sSql & " ORDER BY convert(int, A.cAgeCod)"
'
'    Case "Ahorros"
'
'        If gbBitCentral = True Then
'
'            If lbAhorrosRestringidos Then
'                    lsWhere1 = " (nMoneda=" & pnMoneda & " and nProducto=" & Producto.gCapAhorros & ") "
'                    lsWhere2 = " (Cs.nMoneda=" & pnMoneda & " and CS.nProducto=" & Producto.gCapAhorros & ")"
'
'                    sSql = " Select T.cAgeCod, R.nSaldo"
'                    sSql = sSql & " From ("
'                    sSql = sSql & "     select (Case when cAgeCod='99' then '01' else cAgeCod end) cAgeCod"
'                    sSql = sSql & "     From agencias"
'                    sSql = sSql & "     Group by Case when cAgeCod='99' then '01' else cAgeCod end"
'                    sSql = sSql & " ) T LEFT JOIN ( "
'
'                    sSql = sSql & "Select cAgeCod=Case when A.cAgeCod='99' then '01' else A.cAgeCod end, SUM(A.nSaldo) as nSaldo "
'                    sSql = sSql & " From( "
'                    '
'                    sSql = sSql & " Select CI.cAgeCod, isnull(C3.nSaldo,0)  - isnull((Select Sum(nSaldCnt) From Capsaldosdiarios where dfecha >= '" & Format(ldFecha, "yyyy/mm/dd") & "' And dfecha < '" & Format(DateAdd("d", 1, ldFecha), "yyyy/mm/dd") & "' and (ninmovilizada = 1  or nTpoBloqueo in (3,15) ) and substring(cCtaCod, 4,2) = CI.cAgeCod and substring(cCtaCod, 6,3) = '232' and substring(cCtaCod, 9,1) = '" & pnMoneda & "'),0) as nSaldo From "
'                        'Fecha Guia
'                    sSql = sSql & " ( select right(cCodAge,2) cAgeCod, Convert(Varchar(8),dEstad,112) as dEstad from capestadsaldo " & _
'                        " WHERE " & lsWhere1 & " " & _
'                        " GROUP BY right(cCodAge,2), Convert(Varchar(8),dEstad,112)) CI "
'                    sSql = sSql & " LEFT JOIN ( Select right(cCodAge,2) cAgeCod, Convert(Varchar(8),dEstad,112) as dEstad, " _
'                         & "              sum(nSaldo) nSaldo From CapEstadSaldo " _
'                         & " Where nMoneda = " & pnMoneda & " And nProducto = " & Producto.gCapAhorros _
'                         & " GROUP BY right(cCodAge,2), Convert(Varchar(8),dEstad,112) ) C3 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C3.dEstad,112) AND CI.cAgeCod=C3.cAgeCod "
'                    sSql = sSql & " INNER JOIN Anx10Datos a on right(CI.cAgeCod,2) = A.cAgeCod Right Join " & _
'                           " (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, " & _
'                           " A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod " & _
'                           " ) c2 ON right(C2.cAgeCod,2) = CI.cAgeCod "
'                    sSql = sSql & " WHERE CI.dEstad=( SELECT CONVERT(VARCHAR(8),MAX(CS.dEstad),112)  " & _
'                        " FROM CapEstadSaldo CS " & _
'                        " WHERE convert(varchar(8), CS.dEstad,112) = '" & Format(ldFecha, gsFormatoMovFecha) & "' AND " & lsWhere2 & ")"
'                    sSql = sSql & " ) A GROUP BY Case when A.cAgeCod='99' then '01' else A.cAgeCod end "
'                    sSql = sSql & " ) R ON R.cAgeCod = T.cAgeCod "
'                    sSql = sSql & " Order by T.cAgeCod "
'                Else
'                    lsWhere1 = " (nMoneda=" & pnMoneda & " and nProducto=" & Producto.gCapAhorros & ") "
'                    lsWhere2 = " (Cs.nMoneda=" & pnMoneda & " and CS.nProducto=" & Producto.gCapAhorros & ")"
'
'                    sSql = " Select T.cAgeCod, R.nSaldo"
'                    sSql = sSql & " From ("
'                    sSql = sSql & "     select (Case when cAgeCod='99' then '01' else cAgeCod end) cAgeCod"
'                    sSql = sSql & "     From agencias"
'                    sSql = sSql & "     Group by Case when cAgeCod='99' then '01' else cAgeCod end"
'                    sSql = sSql & " ) T LEFT JOIN ( "
'
'                    sSql = sSql & "Select cAgeCod=Case when A.cAgeCod='99' then '01' else A.cAgeCod end, SUM(A.nSaldo) as nSaldo "
'                    sSql = sSql & " From( "
'                    '
'                    sSql = sSql & " Select CI.cAgeCod, isnull(C3.nSaldo,0) as nSaldo From "
'                        'Fecha Guia
'                    sSql = sSql & " ( select right(cCodAge,2) cAgeCod, Convert(Varchar(8),dEstad,112) as dEstad from capestadsaldo " & _
'                        " WHERE " & lsWhere1 & " " & _
'                        " GROUP BY right(cCodAge,2), Convert(Varchar(8),dEstad,112)) CI "
'                    sSql = sSql & " LEFT JOIN ( Select right(cCodAge,2) cAgeCod, Convert(Varchar(8),dEstad,112) as dEstad, " _
'                         & "              sum(nSaldo) nSaldo From CapEstadSaldo " _
'                         & " Where nMoneda = " & pnMoneda & " And nProducto = " & Producto.gCapAhorros _
'                         & " GROUP BY right(cCodAge,2), Convert(Varchar(8),dEstad,112) ) C3 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C3.dEstad,112) AND CI.cAgeCod=C3.cAgeCod "
'                    sSql = sSql & " INNER JOIN Anx10Datos a on right(CI.cAgeCod,2) = A.cAgeCod Right Join " & _
'                           " (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, " & _
'                           " A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod " & _
'                           " ) c2 ON right(C2.cAgeCod,2) = CI.cAgeCod "
'                    sSql = sSql & " WHERE CI.dEstad=( SELECT CONVERT(VARCHAR(8),MAX(CS.dEstad),112)  " & _
'                        " FROM CapEstadSaldo CS " & _
'                        " WHERE convert(varchar(8), CS.dEstad,112) = '" & Format(ldFecha, gsFormatoMovFecha) & "' AND " & lsWhere2 & ")"
'                    sSql = sSql & " ) A GROUP BY Case when A.cAgeCod='99' then '01' else A.cAgeCod end "
'                    sSql = sSql & " ) R ON R.cAgeCod = T.cAgeCod "
'                    sSql = sSql & " Order by T.cAgeCod "
'                End If
'        Else
'            sSql = " select  c2.cAgeCod, isnull(C1.nSaldo, 0) as nSaldo From ( " & _
'                   " SELECT right(e.cCodAge,2) cCodAge , SUM(e.nSaldPN + e.nSaldPSFL + e.nSaldPCFL + " & _
'                   " e.nSaldCMAC + e.nSaldCRAC + e.nSaldFoncodes + e.nSaldInPN + e.nSaldInPSFL + e.nSaldInPCFL) " & _
'                   " AS nSaldo FROM dbo.Anx10Datos a INNER JOIN " & _
'                   sservidorconsolidada & "EstadMensAho e ON RIGHT(e.cCodAge, 2) = a.cAgeCod " & _
'                   " WHERE dEstadMens = '" & Format(ldFecha, gsFormatoFecha) & "' AND (e.cmoneda =" & pnMoneda & " ) " & _
'                   " GROUP BY e.cCodAge) c1 Right Join " & _
'                   " (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, " & _
'                   " A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod " & _
'                   " ) c2 ON C1.cCodAge = C2.cAgeCod ORDER BY convert(int, C2.cAgeCod)"
'        End If
'    Case "Plazo" 'Plazo Fijo
'         If gbBitCentral = True Then
'            If lbAhorrosRestringidos Then
'                lsWhere1 = " (nMoneda=" & pnMoneda & " and nProducto=" & Producto.gCapPlazoFijo & ") "
'                lsWhere2 = " (Cs.nMoneda=" & pnMoneda & " and CS.nProducto=" & Producto.gCapPlazoFijo & ") "
'
'                lsWhere1C = " (nMoneda=" & pnMoneda & " and nProducto=" & Producto.gCapCTS & ") "
'                lsWhere2C = " (Cs.nMoneda=" & pnMoneda & " and CS.nProducto=" & Producto.gCapCTS & ") "
'
'                '
'                'sSql = " Select T.cAgeCod, R.nSaldo + isnull((Select Sum(nSaldCnt) From Capsaldosdiarios where datediff(day,dfecha,' " & Format(ldFecha, "yyyy/mm/dd") & " ') = 0 and (ninmovilizada = 1 or nCuentaCredMV = 1 or nTpoBloqueo in (3,15) ) and substring(cCtaCod, 4,2) = T.cAgeCod and substring(cCtaCod, 6,3) = '232' and substring(cCtaCod, 9,1) = '" & pnMoneda & "'),0) nSaldo "
'                sSql = " Select T.cAgeCod, R.nSaldo + isnull((Select Sum(nSaldCnt) From Capsaldosdiarios where dfecha >= '" & Format(ldFecha, "yyyy/mm/dd") & "' And dfecha < '" & Format(DateAdd("d", 1, ldFecha), "yyyy/mm/dd") & "' and (ninmovilizada = 1 or nTpoBloqueo in (3,15) ) and substring(cCtaCod, 4,2) = T.cAgeCod and substring(cCtaCod, 6,3) = '232' and substring(cCtaCod, 9,1) = '" & pnMoneda & "'),0) nSaldo "
'                sSql = sSql & " From ("
'                sSql = sSql & "     select (Case when cAgeCod='99' then '01' else cAgeCod end) cAgeCod"
'                sSql = sSql & "     From agencias"
'                sSql = sSql & "     Group by Case when cAgeCod='99' then '01' else cAgeCod end"
'                sSql = sSql & " ) T LEFT JOIN ( "
'
'                sSql = sSql & " Select cAgeCod=Case when A.cAgeCod='99' then '01' else A.cAgeCod end, SUM(A.nSaldo) as nSaldo "
'                sSql = sSql & " From( "
'                '
'                sSql = sSql & "select CI.cAgeCod, isnull(C3.nSaldo,0) as nSaldo From "
'
'                'Fecha Guia
'                sSql = sSql & " ( select right(cCodAge,2) cAgeCod, Convert(Varchar(8),dEstad,112) as dEstad from capestadsaldo " & _
'                    " WHERE (" & lsWhere1 & " or " & lsWhere1C & ") " & _
'                    " GROUP BY right(cCodAge,2), Convert(Varchar(8),dEstad,112)) CI "
'                sSql = sSql & " LEFT JOIN ( Select right(cCodAge,2) cAgeCod, Convert(Varchar(8),dEstad,112) as dEstad, " _
'                     & "              sum(nSaldo) nSaldo From CapEstadSaldo " _
'                     & " Where nMoneda = " & pnMoneda & " And nProducto IN (" & Producto.gCapPlazoFijo & "," & Producto.gCapCTS & ") " _
'                     & " GROUP BY right(cCodAge,2), Convert(Varchar(8),dEstad,112) ) C3 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C3.dEstad,112) AND CI.cAgeCod=C3.cAgeCod "
'                sSql = sSql & " INNER JOIN Anx10Datos a on right(CI.cAgeCod,2) = A.cAgeCod Right Join " & _
'                       " (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, " & _
'                       " A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod " & _
'                       " ) c2 ON right(C2.cAgeCod,2) = CI.cAgeCod "
'                sSql = sSql & " WHERE CI.dEstad=( SELECT CONVERT(VARCHAR(8),MAX(CS.dEstad),112)  " & _
'                    " FROM CapEstadSaldo CS " & _
'                    " WHERE Convert(varchar(8), CS.dEstad,112) = '" & Format(ldFecha, gsFormatoMovFecha) & "'  AND (" & lsWhere2 & " OR " & lsWhere2C & "))"
'                'sSql = sSql & " ORDER BY convert(int, CI.cAgeCod)"
'                sSql = sSql & " ) A GROUP BY Case when A.cAgeCod='99' then '01' else A.cAgeCod end "
'                'sSql = sSql & " ORDER BY Case when A.cAgeCod='02' then '01' else A.cAgeCod end "
'                sSql = sSql & " ) R ON R.cAgeCod = T.cAgeCod "
'                sSql = sSql & " Order by T.cAgeCod "
'            Else
'                lsWhere1 = " (nMoneda=" & pnMoneda & " and nProducto=" & Producto.gCapPlazoFijo & ") "
'                lsWhere2 = " (Cs.nMoneda=" & pnMoneda & " and CS.nProducto=" & Producto.gCapPlazoFijo & ") "
'
'                lsWhere1C = " (nMoneda=" & pnMoneda & " and nProducto=" & Producto.gCapCTS & ") "
'                lsWhere2C = " (Cs.nMoneda=" & pnMoneda & " and CS.nProducto=" & Producto.gCapCTS & ") "
'
'                sSql = " Select T.cAgeCod, R.nSaldo nSaldo "
'                sSql = sSql & " From ("
'                sSql = sSql & "     select (Case when cAgeCod='99' then '01' else cAgeCod end) cAgeCod"
'                sSql = sSql & "     From agencias"
'                sSql = sSql & "     Group by Case when cAgeCod='99' then '01' else cAgeCod end"
'                sSql = sSql & " ) T LEFT JOIN ( "
'
'                sSql = sSql & " Select cAgeCod=Case when A.cAgeCod='99' then '01' else A.cAgeCod end, SUM(A.nSaldo) as nSaldo "
'                sSql = sSql & " From( "
'                '
'                sSql = sSql & "select CI.cAgeCod, isnull(C3.nSaldo,0) as nSaldo From "
'
'                'Fecha Guia
'                sSql = sSql & " ( select right(cCodAge,2) cAgeCod, Convert(Varchar(8),dEstad,112) as dEstad from capestadsaldo " & _
'                    " WHERE (" & lsWhere1 & " or " & lsWhere1C & ") " & _
'                    " GROUP BY right(cCodAge,2), Convert(Varchar(8),dEstad,112)) CI "
'                sSql = sSql & " LEFT JOIN ( Select right(cCodAge,2) cAgeCod, Convert(Varchar(8),dEstad,112) as dEstad, " _
'                     & "              sum(nSaldo) nSaldo From CapEstadSaldo " _
'                     & " Where nMoneda = " & pnMoneda & " And nProducto IN (" & Producto.gCapPlazoFijo & "," & Producto.gCapCTS & ") " _
'                     & " GROUP BY right(cCodAge,2), Convert(Varchar(8),dEstad,112) ) C3 ON Convert(Varchar(8),CI.dEstad,112)=Convert(Varchar(8),C3.dEstad,112) AND CI.cAgeCod=C3.cAgeCod "
'                sSql = sSql & " INNER JOIN Anx10Datos a on right(CI.cAgeCod,2) = A.cAgeCod Right Join " & _
'                       " (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, " & _
'                       " A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod " & _
'                       " ) c2 ON right(C2.cAgeCod,2) = CI.cAgeCod "
'                sSql = sSql & " WHERE CI.dEstad=( SELECT CONVERT(VARCHAR(8),MAX(CS.dEstad),112)  " & _
'                    " FROM CapEstadSaldo CS " & _
'                    " WHERE Convert(varchar(8), CS.dEstad,112) = '" & Format(ldFecha, gsFormatoMovFecha) & "'  AND (" & lsWhere2 & " OR " & lsWhere2C & "))"
'                sSql = sSql & " ) A GROUP BY Case when A.cAgeCod='99' then '01' else A.cAgeCod end "
'                sSql = sSql & " ) R ON R.cAgeCod = T.cAgeCod "
'                sSql = sSql & " Order by T.cAgeCod "
'            End If
'         Else
'            sSql = " select  C2.cAgeCod, isnull(C1.nSaldo, 0) as nSaldo From ( " _
'                    & " select right(pf.cCodAge,2) cCodAge, Sum(nSaldOrd+nSaldConv+nSaldPN+nSaldPSFL+nSaldPCFL+nSaldCmac+nSaldCrac) nSaldo " _
'                    & "from Anx10Datos a JOIN " _
'                    & "   " & sservidorconsolidada & "EstadMensPF pf ON RIGHT(pf.cCodAge,2) = a.cAgeCod left JOIN " _
'                    & "   " & sservidorconsolidada & "EstadMensCTS cts ON pf.dEstadMens = cts.dEstadMens and pf.cCodAge = cts.cCodAge and pf.cmoneda = cts.cmoneda " _
'                    & "where datediff(d,pf.dEstadMens,'" & Format(ldFecha, gsFormatoFecha) & "') = 0 " _
'                    & "and pf.cmoneda = " & pnMoneda & " " _
'                    & "Group by pf.cCodAge ) c1 Right Join " _
'                    & " (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, " _
'                    & " A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod " _
'                    & " ) c2 ON C1.cCodAge = C2.cAgeCod ORDER BY convert(int, C2.cAgeCod)"
'         End If
'
'     Case "Creditos" 'Colocaciones
'         If gsCodCMAC = "108" Then  'La Caja de Lima no tiene la 14 a nivel de Agencia
'            sSql = "SELECT anx.cAgeCod, isnull(cr.nSaldo,0) + isnull(pig.nSaldo,0) nSaldo " _
'                 & "FROM anx10datos anx JOIN Agencias a on anx.cAgeCod = a.cAgeCod LEFT JOIN "
'
'            'sSql = sSql & " (SELECT cCodAge cAgeCod, IsNull(SUM(nSaldoCap),0) as nSaldo "
'
'            sSql = sSql & " (SELECT cAgeCod = case when cCodAge='99' then '01' else cCodAge end, IsNull(SUM(nSaldoCap),0) as nSaldo "
'            sSql = sSql & "  From ColocEstadDiaCred "
'            sSql = sSql & "   WHERE  DATEDIFF(d,dEstad, '" & Format(ldFecha, gsFormatoFecha) & "') = 0    and SubString(cLineaCred,5,1)='" & pnMoneda & "' "
'            'sSql = sSql & "   group by cCodAge "
'            sSql = sSql & "   group by case when cCodAge='99' then '01' else cCodAge end "
'            sSql = sSql & "  ) Cr ON Cr.cAgeCod = anx.cAgeCod LEFT JOIN "
'            'sSql = sSql & "   (SELECT cCodAge cAgeCod, "
'            sSql = sSql & "   (SELECT cAgeCod = case when right(cCodAge,2)='99' then '01' else right(cCodAge,2) end, "
'            sSql = sSql & " IsNull(SUM(nCapVig),0) as nSaldo "
'            sSql = sSql & "    From ColocEstadDiaPrenda "
'            sSql = sSql & "    WHERE  DATEDIFF(d,dEstad, '" & Format(ldFecha, gsFormatoFecha) & "') = 0    and SubString(cLineaCred,5,1)='" & pnMoneda & "' "
'            sSql = sSql & "    group by case when right(cCodAge,2)='99' then '01' else right(cCodAge,2) end "
'            sSql = sSql & "   ) Pig ON Pig.cAgeCod = anx.cAgeCod "
'            sSql = sSql & " Where anx.bcambiarage=0 Order By anx.cAgeCod"
'
'         Else
'            sSql = " select  C2.cAgeCod, isnull(C1.nSaldo, 0) as nSaldo From ( " _
'                    & "SELECT right(cCtaContCod,2) as cCodAge, CASE WHEN right(cCtaContCod,2) = '01' THEN '01' " _
'                    & "   WHEN ag.nAgeEspecial = 1 THEN '04' ELSE '03' END cTipo, Sum(nCtaSaldoImporte) nSaldo " _
'                    & "from CtaSaldo cs join agencias ag ON ag.cAgeCod = right(cCtaContCod,2) " _
'                    & "where cs.cCtaContCod like '14" & Trim(IIf(pnMoneda = 1, "[13]", pnMoneda)) & "[1456]%' and cs.dCtaSaldoFecha = (Select Max(dCtaSaldoFecha) FROM CtaSaldo cs1 " _
'                    & "      where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & Format(ldFecha, gsFormatoFecha) & "') " _
'                    & "group by right(cCtaContCod,2), CASE WHEN right(cCtaContCod,2) = '01' THEN '01' " _
'                    & "   WHEN ag.nAgeEspecial = 1 THEN '04' ELSE '03' END ) c1 Right Join " _
'                    & " (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, " _
'                    & " A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod " _
'                    & " ) c2 ON C1.cCodAge = C2.cAgeCod ORDER BY convert(int, C2.cAgeCod)"
'        End If
'
'    Case "Gerentes"
'        'Se modifico la consulta por que esta sacando gerentes encargados.
'        'By GITU 13-01-2009
'        If gsCodCMAC = "102" Then 'ok
'             sSql = " select  C2.cAgeCod, isnull(C1.nSaldo, 0) as nSaldo From ( " _
'                    & " Select Count(*) nSaldo, RHC.cRHAgenciaCod from RHCargos RHC " _
'                    & "    Inner Join RRHH RH ON RH.cPersCod = RHC.cPersCod " _
'                    & "Where dRHCargoFecha = (Select Max(dRHCargoFecha) From RHCargos RHC1 Where dRHCargoFecha < '" & Format(ldFecha, gsFormatoFecha) & "' " _
'                    & "            And RHC.cPersCod = RHC1.cPersCod) " _
'                    & "And Left(nRHEstado,1) Not In ('7','8') And Left(RHC.cRHCargoCod,3) IN ('002','003','004') And RH.cRHCod Like 'E%' Group By RHC.cRHAgenciaCod) c1 Right Join " _
'                    & " (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, " _
'                    & " A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod " _
'                    & " ) c2 ON C1.cRHAgenciaCod = C2.cAgeCod ORDER BY convert(int, C2.cAgeCod)"
'        Else
'            sSql = " select  C2.cAgeCod, isnull(C1.nSaldo, 0) as nSaldo From ( "
'            'sSql = sSql & " Select Count(*) nSaldo, RHC.cRHAgenciaCod "
'            sSql = sSql & " Select Count(*) nSaldo, cRHAgenciaCod=case when RHC.cRHAgenciaCod ='02' then '01' else RHC.cRHAgenciaCod end "
'            sSql = sSql & " from RHCargos RHC "
'            sSql = sSql & "    Inner Join RRHH RH ON RH.cPersCod = RHC.cPersCod "
'            sSql = sSql & " Where dRHCargoFecha = (Select Max(dRHCargoFecha) From RHCargos RHC1 Where dRHCargoFecha < '" & Format(ldFecha, gsFormatoFecha) & "' "
'            sSql = sSql & "            And RHC.cPersCod = RHC1.cPersCod) "
'            sSql = sSql & " And Left(nRHEstado,1) Not In ('7','8') And Left(RHC.cRHCargoCod,3) = '001' and Not Right(RHC.cRHCargoCod,1) = '5' And RH.cRHCod Like 'E%' "
'            'sSql = sSql & " Group By RHC.cRHAgenciaCod) c1 "
'            sSql = sSql & " Group By case when RHC.cRHAgenciaCod ='02' then '01' else RHC.cRHAgenciaCod end) c1 "
'            sSql = sSql & " Right Join (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, "
'            sSql = sSql & " A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod "
'            sSql = sSql & " Where A10.bcambiarAge=0 " '
'            sSql = sSql & " ) c2 ON C1.cRHAgenciaCod = C2.cAgeCod ORDER BY convert(int, C2.cAgeCod)"
'        End If
'
'    Case "Funcionarios" 'ok
'        'sSql = "Select Count(*) nSaldo, RHC.cRHAgenciaCod from RHCargos RHC Inner Join RRHH RH ON RH.cPersCod = RHC.cPersCod Where dRHCargoFecha = (Select Max(dRHCargoFecha) From RHCargos RHC1 Where dRHCargoFecha < '" & Format(ldFecha, gsFormatoFecha) & "' And RHC.cPersCod = RHC1.cPersCod) And Left(nRHEstado,1) Not In ('7','8') And Left(RHC.cRHCargoCod,3) In ('002','003') And RHC.cRHCargoCod <> '002017' And RH.cRHCod Like 'E%' Group By RHC.cRHAgenciaCod
'        'No se hace diferencia entre consolidada y distribuida porque son iguales
'        If gsCodCMAC = "102" Then
'            sSql = " select  C2.cAgeCod, isnull(C1.nSaldo, 0) as nSaldo From ( " _
'                            & " Select Count(*) nSaldo, RHC.cRHAgenciaCod from RHCargos RHC " _
'                            & "Inner Join RRHH RH ON RH.cPersCod = RHC.cPersCod " _
'                            & "Where dRHCargoFecha = (Select Max(dRHCargoFecha) From RHCargos RHC1 Where dRHCargoFecha < '" & Format(ldFecha, gsFormatoFecha) & "' " _
'                            & "            And RHC.cPersCod = RHC1.cPersCod) " _
'                            & "And Left(nRHEstado,1) Not In ('7','8') And Left(RHC.cRHCargoCod,3) In ('005') And RH.cRHCod Like 'E%' Group By RHC.cRHAgenciaCod) c1 Right Join " _
'                            & " (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, " _
'                            & " A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod " _
'                            & " ) c2 ON C1.cRHAgenciaCod = C2.cAgeCod ORDER BY convert(int, C2.cAgeCod)"
'        Else
'            sSql = " select  C2.cAgeCod, isnull(C1.nSaldo, 0) as nSaldo From ( "
'            sSql = sSql & " Select Count(*) nSaldo, "
'            'ssql=ssql & " RHC.cRHAgenciaCod
'            sSql = sSql & " cRHAgenciaCod=case when RHC.cRHAgenciaCod ='02' then '01' else RHC.cRHAgenciaCod end "
'            sSql = sSql & " From RHCargos RHC "
'            sSql = sSql & " Inner Join RRHH RH ON RH.cPersCod = RHC.cPersCod "
'            sSql = sSql & " Where dRHCargoFecha = (Select Max(dRHCargoFecha) From RHCargos RHC1 Where dRHCargoFecha < '" & Format(ldFecha, gsFormatoFecha) & "' "
'            sSql = sSql & "            And RHC.cPersCod = RHC1.cPersCod) "
'            sSql = sSql & " And Left(nRHEstado,1) Not In ('7','8') And Left(RHC.cRHCargoCod,3) In ('002','003') And RHC.cRHCargoCod <> '002017' And RH.cRHCod Like 'E%' "
'            'sSql = sSql & " Group By RHC.cRHAgenciaCod) c1 "
'            sSql = sSql & " Group By case when RHC.cRHAgenciaCod ='02' then '01' else RHC.cRHAgenciaCod end) c1 "
'            sSql = sSql & " Right Join (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, "
'            sSql = sSql & " A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod "
'            sSql = sSql & " Where A10.bcambiarAge=0 " '
'            sSql = sSql & " ) c2 ON C1.cRHAgenciaCod = C2.cAgeCod ORDER BY convert(int, C2.cAgeCod) "
'        End If
'    Case "Empleados"
'        'sSql = "Select Count(*) nSaldo, RHC.cRHAgenciaCod from RHCargos RHC Inner Join RRHH RH ON RH.cPersCod = RHC.cPersCod Where dRHCargoFecha = (Select Max(dRHCargoFecha) From RHCargos RHC1 Where dRHCargoFecha < '" & Format(ldFecha, gsFormatoFecha) & "' And RHC.cPersCod = RHC1.cPersCod) And Left(nRHEstado,1) Not In ('7','8') And Left(RHC.cRHCargoCod,3) Not In ('001','002','003') And RH.cRHCod Like 'E%' Group By RHC.cRHAgenciaCod
'        'No se hace diferencia entre consolidada y distribuida porque son iguales
'        If gsCodCMAC = "102" Then
'            sSql = " select  C2.cAgeCod, isnull(C1.nSaldo, 0) as nSaldo From ( " _
'                    & "Select Count(*) nSaldo, RHC.cRHAgenciaCod from RHCargos RHC " _
'                    & "Inner Join RRHH RH ON RH.cPersCod = RHC.cPersCod " _
'                    & "Inner JOIN RHContrato Cont ON Cont.cPersCod = RH.cPersCod " _
'                    & "      " _
'                    & "Where dRHCargoFecha = (Select Max(dRHCargoFecha) From RHCargos RHC1 Where dRHCargoFecha < '" & Format(ldFecha, gsFormatoFecha) & "' " _
'                    & "            And RHC.cPersCod = RHC1.cPersCod) " _
'                    & "   And nRHContratoTpo = 0 and Cont.cRHContratoNro = (SELECT Max(cRHContratoNro) FROM RHContrato Cont1 WHERE Cont1.cPersCod = Cont.cPersCod ) " _
'                    & "   And Left(nRHEstado,1) Not In ('7','8') And RH.cRHCod Like 'E%' Group By RHC.cRHAgenciaCod) c1 Right Join " _
'                    & " (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, " _
'                    & " A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod " _
'                    & " ) c2 ON C1.cRHAgenciaCod = C2.cAgeCod ORDER BY convert(int, C2.cAgeCod)"
'
'                    '& "   And Left(nRHEstado,1) Not In ('7','8') And Left(RHC.cRHCargoCod,3) Not In ('001','002','003','004','005') And RH.cRHCod Like 'E%' Group By RHC.cRHAgenciaCod) c1 Right Join "
'        Else
'            sSql = " select  C2.cAgeCod, isnull(C1.nSaldo, 0) as nSaldo From ( "
'            sSql = sSql & " Select Count(*) nSaldo, "
'            'sSql = sSql & " RHC.cRHAgenciaCod "
'            sSql = sSql & " cRHAgenciaCod=case when RHC.cRHAgenciaCod ='02' then '01' else RHC.cRHAgenciaCod end "
'            sSql = sSql & " From RHCargos RHC "
'            sSql = sSql & " Inner Join RRHH RH ON RH.cPersCod = RHC.cPersCod "
'            sSql = sSql & " Where dRHCargoFecha = (Select Max(dRHCargoFecha) From RHCargos RHC1 Where dRHCargoFecha < '" & Format(ldFecha, gsFormatoFecha) & "' "
'            sSql = sSql & "             And RHC.cPersCod = RHC1.cPersCod) "
'            sSql = sSql & " And Left(nRHEstado,1) Not In ('7','8')  And RH.cRHCod Like 'E%' And substring(cRHCargoCodOficial,1,3) not in ('001','002','003')"
'            'sSql = sSql & " Group By RHC.cRHAgenciaCod) c1 "
'            sSql = sSql & " Group By case when RHC.cRHAgenciaCod ='02' then '01' else RHC.cRHAgenciaCod end) c1 "
'            sSql = sSql & " Right Join  (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, "
'            sSql = sSql & "  A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod "
'            sSql = sSql & " Where A10.bcambiarAge=0 " '
'            sSql = sSql & " ) c2 ON C1.cRHAgenciaCod = C2.cAgeCod ORDER BY convert(int, C2.cAgeCod) "
'
'            'sSql = sSql & " And Left(nRHEstado,1) Not In ('7','8') And Left(RHC.cRHCargoCod,3) Not In ('001','002','003','004','005') And RH.cRHCod Like 'E%' "
'
'        End If
'    Case "Otros"
'        'sSql = "Select Count(*) nSaldo, RHC.cRHAgenciaCod from RHCargos RHC Inner Join RRHH RH ON RH.cPersCod = RHC.cPersCod Where dRHCargoFecha = (Select Max(dRHCargoFecha) From RHCargos RHC1 Where dRHCargoFecha < '" & Format(ldFecha, gsFormatoFecha) & "' And RHC.cPersCod = RHC1.cPersCod) And Left(nRHEstado,1) Not In ('7','8') And RH.cRHCod Like 'L%' Group By RHC.cRHAgenciaCod "
'         'No se hace diferencia entre consolidada y distribuida porque son iguales
'        If gsCodCMAC = "102" Then
'            sSql = " select  C2.cAgeCod, isnull(C1.nSaldo, 0) as nSaldo From ( " _
'                    & "Select Count(*) nSaldo, RHC.cRHAgenciaCod from RHCargos RHC " _
'                    & "Inner Join RRHH RH ON RH.cPersCod = RHC.cPersCod " _
'                    & "Inner JOIN RHContrato Cont ON Cont.cPersCod = RH.cPersCod " _
'                    & "      " _
'                    & "Where dRHCargoFecha = (Select Max(dRHCargoFecha) From RHCargos RHC1 Where dRHCargoFecha < '" & Format(ldFecha, gsFormatoFecha) & "' " _
'                    & "            And RHC.cPersCod = RHC1.cPersCod) " _
'                    & "   And nRHContratoTpo <> 0 and Cont.cRHContratoNro = (SELECT Max(cRHContratoNro) FROM RHContrato Cont1 WHERE Cont1.cPersCod = Cont.cPersCod ) " _
'                    & "   And Left(nRHEstado,1) Not In ('7','8') And RH.cRHCod LIKE 'E%' and Left(RHC.cRHCargoCod,3) Not In ('001','002','003') Group By RHC.cRHAgenciaCod ) c1 Right Join " _
'                    & " (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, " _
'                    & " A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod " _
'                    & " ) c2 ON C1.cRHAgenciaCod = C2.cAgeCod ORDER BY convert(int, C2.cAgeCod)"
'        Else
'            sSql = " select  C2.cAgeCod, isnull(C1.nSaldo, 0) as nSaldo From ( "
'            sSql = sSql & " Select Count(*) nSaldo, "
'            'sSql = sSql & " RHC.cRHAgenciaCod "
'            sSql = sSql & " cRHAgenciaCod=case when RHC.cRHAgenciaCod ='02' then '01' else RHC.cRHAgenciaCod end "
'            sSql = sSql & " from RHCargos RHC "
'            sSql = sSql & " Inner Join RRHH RH ON RH.cPersCod = RHC.cPersCod "
'            sSql = sSql & " Where dRHCargoFecha = (Select Max(dRHCargoFecha) From RHCargos RHC1 Where dRHCargoFecha < '" & Format(ldFecha, gsFormatoFecha) & "' "
'            sSql = sSql & "            And RHC.cPersCod = RHC1.cPersCod) "
'            sSql = sSql & " And Left(nRHEstado,1) Not In ('7','8') And RH.cRHCod Like '[P]%' "
'            'sSql = sSql & " Group By RHC.cRHAgenciaCod ) c1 "
'            sSql = sSql & " Group By case when RHC.cRHAgenciaCod ='02' then '01' else RHC.cRHAgenciaCod end) c1 "
'            sSql = sSql & " Right Join (select A.cAgeCod, A.cAgeDescripcion, A10.cPais, A10.cDepa, A10.cProv, A10.cDist, "
'            sSql = sSql & "  A10.cTipo, A10.cCodigo From agencias A left join anx10datos A10 on A.cAgeCod=A10.cAgeCod "
'            sSql = sSql & " Where A10.bcambiarAge=0 " '
'            sSql = sSql & "  ) c2 ON C1.cRHAgenciaCod = C2.cAgeCod ORDER BY convert(int, C2.cAgeCod) "
'         End If
'    End Select
'JOEP20190531 Mejora Anexo10

    Set CargaDatosAnexo = oCon.CargaRecordSet(sSql)
End Function

Private Sub cmdGrabar_Click()
Dim clsAnx10 As DAnx10Datos
Dim I As Integer
On Error GoTo AceptarErr
If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox(" ¿ Está seguro de grabar los datos ?      ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
    Select Case opeGraba
    Case 1
          'Nuevo
    Case 2
          Set clsAnx10 = New DAnx10Datos
          For I = 1 To fg.Rows - 1
              clsAnx10.ActualizaAnx10Datos fg.TextMatrix(I, 8), fg.TextMatrix(I, 2), fg.TextMatrix(I, 3), fg.TextMatrix(I, 4), fg.TextMatrix(I, 5), "0" & Trim(Right(fg.TextMatrix(I, 6), 1)), fg.TextMatrix(I, 7)
          Next
          Set clsAnx10 = Nothing
    End Select
    ActivaBotonesEditar False
    CargaDatosUbigeo
    MsgBox "Datos Grabados satisfactoriamente", vbInformation, "Aviso!"
    fg.SetFocus
End If
Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Function ValidaDatos() As Boolean
Dim I As Integer
    ValidaDatos = True
    
    For I = 1 To fg.Rows - 1
        
        If Len(Trim(fg.TextMatrix(I, 2))) = 0 Or Len(Trim(fg.TextMatrix(I, 2))) > 4 Then
            MsgBox "Ud. debe ingresar un código de país válido para " & fg.TextMatrix(I, 1), vbExclamation, "Aviso!!!"
            fg.SetFocus
            ValidaDatos = False
            Exit For
        Else
            If Len(Trim(fg.TextMatrix(I, 3))) = 0 Or Len(Trim(fg.TextMatrix(I, 3))) > 2 Then
                MsgBox "Ud. debe ingresar un código de departamento válido para " & fg.TextMatrix(I, 1), vbExclamation, "Aviso!!!"
                fg.SetFocus
                ValidaDatos = False
                Exit For
            Else
                If Len(Trim(fg.TextMatrix(I, 4))) = 0 Or Len(Trim(fg.TextMatrix(I, 4))) > 2 Then
                    MsgBox "Ud. debe ingresar un código de provincia válido para " & fg.TextMatrix(I, 1), vbExclamation, "Aviso!!!"
                    fg.SetFocus
                    ValidaDatos = False
                    Exit For
                Else
                    If Len(Trim(fg.TextMatrix(I, 5))) = 0 Or Len(Trim(fg.TextMatrix(I, 5))) > 2 Then
                        MsgBox "Ud. debe ingresar un código de distrito válido para " & fg.TextMatrix(I, 1), vbExclamation, "Aviso!!!"
                        fg.SetFocus
                        ValidaDatos = False
                        Exit For
                    Else
                        If Len(Trim(fg.TextMatrix(I, 6))) = 0 Then
                            MsgBox "Ud. debe seleccionar un tipo de oficina válido para " & fg.TextMatrix(I, 1), vbExclamation, "Aviso!!!"
                            fg.SetFocus
                            ValidaDatos = False
                            Exit For
                        Else
                            If Len(Trim(fg.TextMatrix(I, 7))) = 0 Or Len(Trim(fg.TextMatrix(I, 7))) > 3 Then
                                MsgBox "El valor del código no es válido para " & fg.TextMatrix(I, 1), vbExclamation, "Aviso!!!"
                                fg.SetFocus
                                ValidaDatos = False
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
          
End Function

Private Sub cmdNuevo_Click()
opeGraba = 1
'ActivaBotonesEditar True
'fg.AdicionaFila
'fg.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Set oCon = New DConecta
oCon.AbreConexion
CargaDatosUbigeo
cboMes.ListIndex = Month(gdFecSis) - 1
txtAnio = Year(gdFecSis)

Dim oCons As New DConstantes
fg.CargaCombo oCons.RecuperaConstantes(gAnx10TipoOficina)
End Sub

Private Sub CargaDatosUbigeo()
    sSql = "Select cAgeDescripcion, ISNULL(an.cPais,'') cPais, ISNULL(an.cDepa,'') cDepa, " _
        & "       ISNULL(an.cProv,'') cProv, ISNULL(an.cDist,'') cDist, " _
        & "       LEFT( ISNULL(c.cConsDescripcion,'') + Space(75), 75) + ISNULL(Convert(varchar(3),nConsCod),'')  cConsDescripcion, ISNULL(an.cCodigo,'') cCodigo, ag.cAgeCod, ISNULL(an.cTipo,'') cTipo " _
        & "From Agencias ag LEFT JOIN Anx10Datos an ON an.cAgeCod = ag.cAgeCod " _
        & "     LEFT JOIN Constante c ON c.nConsValor = convert(int,an.cTipo) and c.nConsCod = " & gAnx10TipoOficina
    
    Set fg.Recordset = oCon.CargaRecordSet(sSql)
End Sub

Private Sub Form_Unload(Cancel As Integer)
oCon.CierraConexion
Set oCon = Nothing
End Sub

Private Sub txtAnio_GotFocus()
fEnfoque txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    cmdGenerar.SetFocus
End If
End Sub

