VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReporteKardexActivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kardex de Activos"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "frmReporteKardexActivo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Serie:"
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
      Height          =   990
      Left            =   5400
      TabIndex        =   15
      Top             =   0
      Width           =   3060
      Begin Sicmact.TxtBuscar txtSerie 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
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
         TipoBusqueda    =   2
      End
      Begin VB.Label lblSerieG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2835
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rango:"
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
      Height          =   990
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   2100
      Begin MSComCtl2.DTPicker txtFechaIni 
         Height          =   315
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   74186753
         CurrentDate     =   41414
      End
      Begin MSComCtl2.DTPicker txtFechaFin 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Top             =   600
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   74186753
         CurrentDate     =   41414
      End
      Begin VB.Label Label5 
         Caption         =   "Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   280
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   645
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "&Buscar"
      Height          =   325
      Left            =   8520
      TabIndex        =   4
      Top             =   360
      Width           =   1000
   End
   Begin VB.CommandButton cmdImpConsol 
      Caption         =   "&Consolidado"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4260
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3195
      Left            =   90
      TabIndex        =   10
      Top             =   975
      Width           =   9405
      Begin MSComctlLib.ListView lvwMovKardex 
         Height          =   2865
         Left            =   120
         TabIndex        =   5
         Top             =   195
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5054
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Detalle Movimiento"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "REI"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Saldo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Depreciacion"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Valor"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Depreciacion Acu."
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Valor Neto "
            Object.Width           =   1940
         EndProperty
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8250
      TabIndex        =   8
      Top             =   4260
      Width           =   1215
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   7000
      TabIndex        =   7
      Top             =   4260
      Width           =   1215
   End
   Begin VB.Frame fraRango 
      Caption         =   "Bien:"
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
      Height          =   990
      Left            =   2280
      TabIndex        =   9
      Top             =   0
      Width           =   3060
      Begin Sicmact.TxtBuscar txtBS 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
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
      End
      Begin VB.Label lblBienG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2835
      End
   End
End
Attribute VB_Name = "frmReporteKardexActivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim lbLibroOpen As Boolean
Dim xlHoja1 As Excel.Worksheet
Dim nLin As Long
Dim lsArchivo As String

Private Sub cmdBuscar_Click()
Dim oALmacen As DLogAlmacen
Dim oBien As DBien
Dim rs As ADODB.Recordset
Dim lista As ListItem

Set rs = New ADODB.Recordset
Set oALmacen = New DLogAlmacen

'   If txtBS.Text = "" Then
'      MsgBox "Ud. debe seleccionar la Categoría del Bien", vbInformation, "Aviso"
'      txtBS.SetFocus
'      Exit Sub
'   End If
   If txtSerie.Text = "" Then
      MsgBox "Ud. debe seleccionar una Serie para ver su detalle", vbInformation, "Aviso"
      txtSerie.SetFocus
      Exit Sub
   End If
   
   On Error GoTo ErrBuscar
   Screen.MousePointer = 11
   Set oBien = New DBien
   'Set rs = oALmacen.GetKardexActivo(Trim(txtBS.Text), Trim(txtSerie.Text)) 'Falta la Fecha
   Set rs = oBien.GetAFDepreciacionxKardex(Trim(txtSerie.Text))
   Set oALmacen = Nothing

    lvwMovKardex.ListItems.Clear
    Dim J As Integer
    J = 1
    If Not (rs.EOF And rs.BOF) Then
        Do Until rs.EOF
            Set lista = lvwMovKardex.ListItems.Add(, , J)
            lista.SubItems(1) = rs(0) 'Fecha
            lista.SubItems(2) = rs(2) 'Descripcion
            lista.SubItems(3) = rs(3) 'REI
            lista.SubItems(4) = rs(4) 'Saldo
            lista.SubItems(5) = rs(5) 'Depreciacion
            'lista.SubItems(6) = rs(6) 'Ajuste
            'lista.SubItems(7) = rs(7) 'Depreciacion Acumulada
            'lista.SubItems(7) = rs(8) 'Valor Neto en Libro
            rs.MoveNext
            J = J + 1
        Loop
    Else
        
        MsgBox "No existen Datos", vbInformation, "Aviso"
    End If
    Set oBien = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrBuscar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdImpConsol_Click()
Dim oALmacen As DLogAlmacen
Dim rsCon As ADODB.Recordset
Dim lsHoja As String
Dim lsCSerie As String
Dim lsSerie As String
Dim lsDescrip As String
Dim lsArea As String
Dim lsAgencia As String
Dim lsFechaAdq As String
Dim lsCadena As String
Dim lsRuc As String
Dim lsSerieAnt As String
Dim lnVidaUtil As Integer
Dim lsObser As String
Dim lsFechaCond As String
Set rsCon = New ADODB.Recordset
Set oALmacen = New DLogAlmacen
'ALPA 20081222***************************
Dim pnDepreAc As Double
Dim pnValorFin As Double
'****************************************
Dim lsFechaIni As String, lsFechaFin As String
Dim oBien As New DBien
'lsArchivo = App.path & "\SPOOLER\RSARViaticos_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
'
'lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
'If Not lbLibroOpen Then
'    Exit Sub
'End If

'lsFechaCond = Format(mskFI.Text, "yyyymmaa")
lsFechaIni = Format(CDate(txtFechaIni.value), "yyyymmdd")
lsFechaFin = Format(CDate(txtFechaFin.value), "yyyymmdd")

'Set rsCon = oALmacen.GetKardexActivoConsol(lsFechaCond)
Set rsCon = oBien.GetAFDepreciacionxKardexConsol(lsFechaIni, lsFechaFin) 'EJVG20130705

Set oALmacen = Nothing
Set oBien = Nothing
lsRuc = LeeConstanteSist(gConstSistCMACRuc)
lsCadena = ""
nLin = 1
'lsHoja = "KardexActivos"
'ExcelAddHoja lsHoja, xlLibro, xlHoja1
lsSerieAnt = rsCon!cSerie
Do While Not rsCon.EOF
    lsSerie = rsCon!cSerie
    lsDescrip = rsCon!cDescripcion
    lsArea = rsCon!cAreaDescripcion
    lsAgencia = rsCon!cAgeDescripcion
    'ALPA 20081222**************************************
    'lsFechaAdq = rsCon!dActivacion
    lsFechaAdq = rsCon!dCompra
    'lnVidaUtil = rsCon!nBSPerDeprecia
    '***************************************************
    If lsSerie <> lsSerieAnt Then
        lsCadena = lsCadena & String(105, "_") & Chr(10)
        lsCadena = lsCadena & Space(2) & "VIDA UTIL   " & Space(4) & lnVidaUtil & Chr(10)
        lsCadena = lsCadena & Space(2) & "OBSERVACION " & Space(4) & lsObser & Chr(10)
        lsCadena = lsCadena & String(105, "_") & Chr(10) & Chr(12)
        lsSerieAnt = lsSerie
    End If
    
    If lsCSerie <> lsSerie Then
        'ImprimeActivoFijoCab lsSerie, lsDescrip, "", lsArea, lsAgencia, lsFechaAdq, "", nLin
        lsCadena = lsCadena & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10)
        lsCadena = lsCadena & Space(31) & "CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS S.A. " & Space(10) & "RUC" & Space(1) & lsRuc & Chr(10) & Chr(10)
        lsCadena = lsCadena & Space(43) & "CONTROL DE ACTIVO FIJO" & Chr(10)
        'ALPA 20081222***************
        'lsCadena = lsCadena & Space(41) & "AL" & Space(1) & lsFechaAdq & Chr(10)
        'lsCadena = lsCadena & Space(41) & "AL" & Space(1) & Format(mskFI.Text, "dd/mm/yyyy") & Chr(10)
        lsCadena = lsCadena & Space(41) & "DEL" & Space(1) & Format(CDate(txtFechaIni.value), "dd/mm/yyyy") & " AL " & Format(CDate(txtFechaFin.value), "dd/mm/yyyy") & Chr(10)
        '****************************
        lsCadena = lsCadena & "CODIGO      :" & Space(3) & lsSerie & Chr(10)
        lsCadena = lsCadena & "DESCRIPCION :" & Space(3) & lsDescrip & Chr(10) & Chr(10)
        'lsCadena = lsCadena & "            " & Space(3) & "SERIE " & lsSerie & Chr(10)
        lsCadena = lsCadena & "OFICINA     :" & Space(3) & JIZQ(lsAgencia, 30) & Space(10) & "AREA   :" & lsArea & Chr(10)
        lsCadena = lsCadena & "ADQUISICION :" & Space(3) & lsFechaAdq & Chr(10)
        lsCadena = lsCadena & String(105, "_") & Chr(10)
        lsCadena = lsCadena & Space(4) & "FECHA" & Space(7) & "CONCEPTO" & Space(25) & "DEP MES" & Space(8) & "VALOR ACT" & Space(8) & "DEPRECIACION" & Space(8) & "VALOR" & Chr(10)
        lsCadena = lsCadena & String(105, "_") & Chr(10)
        'ALPA 20081222***********************************
        pnDepreAc = 0
        '************************************************
    End If
    lsCSerie = rsCon!cSerie
    
    'ALPA 20081222*************************************************************
    'lsCadena = lsCadena & Space(3) & Mid(rsCon!cMovNro, 1, 8) & Space(5) & JIZQ(rsCon!cOpeDesc, 24) & Space(8) & JDER(Format(rsCon!Depreciacion, "##,#00.00"), 12) & Space(3) & JDER(Format(rsCon!nBSValor, "##,#00.00"), 12) & Space(3) & JDER(Format(rsCon!nBSValor, "##,#00.00"), 12) & Space(3) & JDER(Format(rsCon!nBSValor, "##,#00.00"), 12) & Space(3) & Chr(10)
    pnDepreAc = pnDepreAc + rsCon!Depreciacion
    pnValorFin = rsCon!nBSValor - pnDepreAc
    If pnValorFin = 0 Then
        pnValorFin = 1
    End If
    lsCadena = lsCadena & Space(3)
    lsCadena = lsCadena & Mid(rsCon!cMovNro, 1, 8) & Space(5)
    lsCadena = lsCadena & JIZQ(rsCon!cOpeDesc, 24) & Space(8)
    lsCadena = lsCadena & JDER(Format(rsCon!Depreciacion, "##,#00.00"), 12) & Space(3)
    lsCadena = lsCadena & JDER(Format(rsCon!nBSValor, "##,#00.00"), 12) & Space(3)
    lsCadena = lsCadena & JDER(Format(pnDepreAc, "##,#00.00"), 12) & Space(3)
    lsCadena = lsCadena & JDER(Format(pnValorFin, "##,#00.00"), 12) & Space(3) & Chr(10)
    lnVidaUtil = rsCon!nBSPerDeprecia
    '**************************************************************************
    'If lsCSerie <> lsSerie Then
    'End If
    
'    xlHoja1.Cells(nLin, 3) =
'    xlHoja1.Cells(nLin, 4) =
'    xlHoja1.Cells(nLin, 5) = rs!nBSValor
'    xlHoja1.Cells(nLin, 6) = rs!nBSValor
          
    rsCon.MoveNext
    If rsCon.EOF Then
        Exit Do
    End If
Loop

Dim MSWord As Word.Application
'Dim MSWordSource As Word.Application
Set MSWord = New Word.Application
'Set MSWordSource = New Word.Application
Dim RangeSource As Word.Range
                
'MSWordSource.Documents.Open FileName:=App.path & "\SPOOLER\Boletas_Pago.doc"

'Set RangeSource = MSWordSource.ActiveDocument.Content
'Lo carga en Memoria
'MSWordSource.ActiveDocument.Content.Copy
'MSWordSource.ActiveDocument
'Crea Nuevo Documento
MSWord.Documents.Add
                
MSWord.Application.Selection.TypeParagraph
'MSWord.Application.Selection.Paste
MSWord.Application.Selection.InsertBreak
                
'MSWordSource.ActiveDocument.Close
'MSWordSource.ActiveDocument.Close
'Set MSWordSource = Nothing
                    
MSWord.Selection.SetRange start:=MSWord.Selection.start, End:=MSWord.ActiveDocument.Content.End
MSWord.Selection.MoveEnd
            
               
MSWord.ActiveDocument.Range.InsertBefore lsCadena
MSWord.ActiveDocument.Select
MSWord.ActiveDocument.Range.Font.Name = "Courier New"
MSWord.ActiveDocument.Range.Font.Size = 8
MSWord.ActiveDocument.Range.Paragraphs.Space1
                
MSWord.Selection.Find.Execute Replace:=wdReplaceAll
MSWord.ActiveDocument.PageSetup.Orientation = wdOrientPortrait
                
MSWord.ActiveDocument.PageSetup.TopMargin = CentimetersToPoints(2)
MSWord.ActiveDocument.PageSetup.LeftMargin = CentimetersToPoints(1)
MSWord.ActiveDocument.PageSetup.RightMargin = CentimetersToPoints(1)

'Documento.PageSetup.RightMargin = CentimetersToPoints(0.5)
                
MSWord.ActiveDocument.SaveAs App.path & "\SPOOLER\Activos_Fijos_" & gsCodUser & Format(Now, "yyyymmsshhmmss") & ".doc"
MSWord.Visible = True
Set MSWord = Nothing

'ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
'CargaArchivo lsArchivo, App.path & "\SPOOLER\"

End Sub

Private Sub CmdImprimir_Click()
Dim rs As ADODB.Recordset

Dim fs              As Scripting.FileSystemObject
Dim xlAplicacion    As Excel.Application
Dim xlLibro         As Excel.Workbook
Dim xlHoja1         As Excel.Worksheet
Dim lbExisteHoja    As Boolean
Dim liLineas        As Integer
Dim i               As Integer
Dim glsArchivo      As String
Dim lsNomHoja       As String

Dim lsFecha As String
Dim lsDocumento As String
Dim lsIngreso As String
Dim lsSalida As String
Dim lsSaldo As String
Dim lsArea As String
Dim lsAge As String
Dim lsValIngreso As String
Dim lsValSalida As String
Dim lsSaldoMont As String
Dim lsIngTotal As String
Dim lsSalTotal As String
Dim lsProveedor As String

Dim lnIngTot As Currency
Dim lnSalTot As Currency
Dim lnIngTotal As Currency
Dim lnSalTotal As Currency

    
    If Me.lvwMovKardex.ListItems.Count < 1 Then 'flex.TextMatrix(1, 1) = "" Then
        MsgBox "Debe generar previamente el Kardex antes de Exportarlo a EXCEL.", vbInformation, "Aviso"
        Me.CmdBuscar.SetFocus
        Exit Sub
    End If
        
    glsArchivo = "Reporte_KARDEX_Activos" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets.Add

    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    xlHoja1.PageSetup.Orientation = xlLandscape

            lbExisteHoja = False
            lsNomHoja = "KARDEX Activos"
            For Each xlHoja1 In xlLibro.Worksheets
                If xlHoja1.Name = lsNomHoja Then
                    xlHoja1.Activate
                    lbExisteHoja = True
                    Exit For
                End If
            Next
            If lbExisteHoja = False Then
                Set xlHoja1 = xlLibro.Worksheets.Add
                xlHoja1.Name = lsNomHoja
            End If

            xlAplicacion.Range("A1:A1").ColumnWidth = 10
            xlAplicacion.Range("B1:B1").ColumnWidth = 20
            xlAplicacion.Range("c1:c1").ColumnWidth = 10
            xlAplicacion.Range("D1:D1").ColumnWidth = 10
            xlAplicacion.Range("E1:E1").ColumnWidth = 10
            xlAplicacion.Range("F1:F1").ColumnWidth = 10
            xlAplicacion.Range("G1:G1").ColumnWidth = 10
            xlAplicacion.Range("H1:H1").ColumnWidth = 10
            xlAplicacion.Range("I1:I1").ColumnWidth = 10
                    
            xlAplicacion.Range("A1:Z100").Font.Size = 9
       
            xlHoja1.Cells(1, 1) = "CAJA MUNICIPAL MAYNAS"
            xlHoja1.Cells(2, 1) = "Activos Fijos"
            xlHoja1.Cells(3, 7) = "Fecha :" & Format(gdFecSis, "dd/mm/yyyy")
            xlHoja1.Cells(4, 7) = "Hora :"
            xlHoja1.Cells(5, 4) = "KARDEX DE ACTIVO"
            
                      
            xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 3)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(5, 4), xlHoja1.Cells(5, 8)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 3)).Merge True
            xlHoja1.Range(xlHoja1.Cells(5, 4), xlHoja1.Cells(5, 8)).Merge True
           
            xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 10)).Borders(xlEdgeBottom).Weight = xlMedium
            xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 10)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            
            xlHoja1.Cells(6, 1) = "Clase de Activo : " & lblBienG.Caption
            xlHoja1.Cells(7, 1) = "Ubicacion : " '& Trim(Me.lblAlmacenG.Caption)
            xlHoja1.Cells(8, 1) = "Descripcion : " '& Trim(Me.lblBien.Caption)
            xlHoja1.Cells(9, 1) = "Codigo : " & Me.txtBS.Text
            xlHoja1.Cells(10, 1) = "Nro. de Serie : " & Me.txtSerie.Text
            xlHoja1.Cells(11, 1) = "Importe Original : "
            xlHoja1.Cells(12, 1) = "Proveedor : "
            xlHoja1.Cells(13, 1) = "Descripcion : "
            
            xlHoja1.Range(xlHoja1.Cells(13, 1), xlHoja1.Cells(13, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            xlHoja1.Range(xlHoja1.Cells(13, 1), xlHoja1.Cells(13, 10)).Borders(xlEdgeBottom).Weight = xlMedium
            xlHoja1.Range(xlHoja1.Cells(13, 1), xlHoja1.Cells(13, 10)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                                  
            liLineas = 14
            
            xlHoja1.Cells(liLineas, 1) = "Fecha Registro"
            xlHoja1.Cells(liLineas, 2) = "Fecha Registro"
            xlHoja1.Cells(liLineas, 3) = "Detalle de Movimiento"
            xlHoja1.Cells(liLineas, 4) = "REI"
            xlHoja1.Cells(liLineas, 5) = "Saldo"
            '''xlHoja1.Cells(liLineas, 6) = "Depreciacion S/." 'marg ers044-2016
            xlHoja1.Cells(liLineas, 6) = "Depreciacion " & gcPEN_SIMBOLO
            xlHoja1.Cells(liLineas, 7) = "valor"
            xlHoja1.Cells(liLineas, 8) = "Depreciacion Acumulada"
            xlHoja1.Cells(liLineas, 9) = "Valor Neto en Libros"
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 10)).Borders(xlEdgeBottom).Weight = xlMedium
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 10)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            
            
           
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 10)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 10)).VerticalAlignment = xlCenter

            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas + 2, 1)).Merge True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 10)).EntireRow.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 10)).WrapText = True
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 10)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 10)).Borders.LineStyle = 1
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 10)).Interior.Color = RGB(159, 206, 238)
            
                   
            liLineas = liLineas + 1
                 
      Dim lnI As Integer
      'EJVG20130705 ***
'      rs.MoveFirst
'      Do Until rs.EOF
'
'            xlHoja1.Cells(liLineas, 1) = Mid(rs(0), 1, 4) & "-" & Mid(rs(0), 5, 2) 'Fecha
'            xlHoja1.Cells(liLineas, 2) = rs(2) 'Detalle Mov
'            xlHoja1.Cells(liLineas, 3) = rs(3) 'REI
'            xlHoja1.Cells(liLineas, 4) = rs(4) 'Saldo
'            xlHoja1.Cells(liLineas, 5) = rs(5) 'Depreciacion
'            xlHoja1.Cells(liLineas, 6) = rs(5) 'Ajuste
'            xlHoja1.Cells(liLineas, 7) = rs(5) 'Depreciacion Acumulada
'            xlHoja1.Cells(liLineas, 8) = rs(5) 'Valor Neto en Libros
'
'            liLineas = liLineas + 1
'            rs.MoveNext
'       Loop
    For lnI = 1 To lvwMovKardex.ListItems.Count - 1
          xlHoja1.Cells(liLineas, 1) = lvwMovKardex.ListItems.item(lnI).SubItems(1) 'Fecha
          xlHoja1.Cells(liLineas, 2) = lvwMovKardex.ListItems.item(lnI).SubItems(2) 'Detalle Mov
          xlHoja1.Cells(liLineas, 3) = lvwMovKardex.ListItems.item(lnI).SubItems(3) 'REI
          xlHoja1.Cells(liLineas, 4) = lvwMovKardex.ListItems.item(lnI).SubItems(4) 'Saldo
          xlHoja1.Cells(liLineas, 5) = lvwMovKardex.ListItems.item(lnI).SubItems(5) 'Depreciacion
          xlHoja1.Cells(liLineas, 6) = lvwMovKardex.ListItems.item(lnI).SubItems(6) 'Ajuste
          xlHoja1.Cells(liLineas, 7) = lvwMovKardex.ListItems.item(lnI).SubItems(7) 'Depreciacion Acumulada
          xlHoja1.Cells(liLineas, 8) = lvwMovKardex.ListItems.item(lnI).SubItems(8) 'Valor Neto en Libros
                      
          liLineas = liLineas + 1
    Next
    'END EJVG *******
    xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 10)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 10)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 10)).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 10)).Interior.Color = RGB(159, 206, 238)
               

        'ExcelCuadro xlHoja1, 1, 4, 12, liLineas - 1
        
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        ExcelEnd App.path & "\Spooler\" & glsArchivo, xlAplicacion, xlLibro, xlHoja1
    
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        Call CargaArchivo(glsArchivo, App.path & "\SPOOLER\")
 
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

'Private Sub Command1_Click()
'Dim oALmacen As DLogAlmacen
'Dim rs As ADODB.Recordset
'Dim lista As ListItem
'
'Set rs = New ADODB.Recordset
'Set oALmacen = New DLogAlmacen
'
'   Set rs = oALmacen.GetKardexActivoConsol()
'   Set oALmacen = Nothing
'End Sub

Private Sub Form_Load()
Dim oALmacen As DLogAlmacen
Set oALmacen = New DLogAlmacen
Dim oBien As New DBien
'EJVG20130705 ***
CentraForm Me
'Me.txtBS.rs = oALmacen.GetAFBienes
    txtFechaIni.value = Format(gdFecSis, gsFormatoFechaView)
    txtFechaFin.value = Format(gdFecSis, gsFormatoFechaView)
    txtBS.rs = oBien.GetAFBienesFull("")
    txtSerie.rs = oBien.GetAFBienesPaKardex(txtBS.Text, CDate("1900-01-01"), gdFecSis) 'EJVG20130705
    Set oBien = Nothing
'END EJVG *******
End Sub

Private Sub txtBS_EmiteDatos()
    Dim oBien As New DBien
    On Error GoTo ErrBSEmite
    Screen.MousePointer = 11
    lblBienG.Caption = ""
    If txtBS.Text <> "" Then
        Me.lblBienG.Caption = txtBS.psDescripcion
        txtSerie.Text = ""
        lblSerieG.Caption = ""
        'Me.txtSerie.rs = oALmacen.GetAFBSSerie(txtBS.Text, Year(gdFecSis))
        txtSerie.rs = oBien.GetAFBienesPaKardex(txtBS.Text, CDate(txtFechaIni.value), CDate(txtFechaFin.value)) 'EJVG20130705
    End If
    Screen.MousePointer = 0
    Set oBien = Nothing
    Exit Sub
ErrBSEmite:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub ImprimeActivoFijoCab(psCodigo As String, psDescrip As String, psSerie As String, psArea As String, _
                                 psAgencia As String, pdFechaAdq As String, psMarca As String, lnLin As Long)
    nLin = lnLin
    
    xlHoja1.Range("A1:G1").EntireColumn.Font.FontStyle = "Arial"
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.TopMargin = 2
    
    xlHoja1.Range("A1:A1").RowHeight = 17
    xlHoja1.Range("A1:A1").ColumnWidth = 8  'Fecha
    xlHoja1.Range("B1:B1").ColumnWidth = 30 'Concepto
    xlHoja1.Range("C1:C1").ColumnWidth = 12 'Dep Mes
    xlHoja1.Range("D1:D1").ColumnWidth = 12 'Valor Act
    xlHoja1.Range("E1:E1").ColumnWidth = 12 'Depreciacion
    xlHoja1.Range("F1:F1").ColumnWidth = 12 'Valor
    
    xlHoja1.Cells(nLin, 2) = "CAJA MUNICIPAL DE AHORRO DE CREDITO DE MAYNAS"
    xlHoja1.Range("A" & nLin & ":F" & nLin).Merge True
    xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":F" & nLin).HorizontalAlignment = xlHAlignCenter
    nLin = nLin + 3
    xlHoja1.Cells(nLin, 2) = "CONTROL DE ACTIVO FIJO"
    xlHoja1.Range("B" & nLin & ":E" & nLin).Font.Bold = True
    xlHoja1.Range("B" & nLin & ":E" & nLin).Merge True
    nLin = nLin + 1
    xlHoja1.Cells(nLin, 2) = " Al " & pdFechaAdq
    xlHoja1.Range("B" & nLin & ":E" & nLin).Font.Bold = True
    xlHoja1.Range("B" & nLin & ":E" & nLin).Merge True
    xlHoja1.Range("B" & nLin & ":E" & nLin).HorizontalAlignment = xlHAlignCenter
    
    nLin = nLin + 1
    
    xlHoja1.Cells(nLin, 1) = "CODIGO "
    xlHoja1.Cells(nLin, 3) = psCodigo
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & nLin & ":C" & nLin).Font.Bold = True
    nLin = nLin + 1
    
    xlHoja1.Cells(nLin, 1) = "DESCRIPCION"
    xlHoja1.Cells(nLin, 3) = psDescrip
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & nLin & ":C" & nLin).Font.Bold = True
    xlHoja1.Range("C" & nLin & ":C" & nLin).HorizontalAlignment = xlHAlignLeft
    nLin = nLin + 1
    
    xlHoja1.Cells(nLin, 3) = "SERIE:"
    xlHoja1.Cells(nLin, 4) = psSerie
    xlHoja1.Cells(nLin, 6) = "MARCA:"
    xlHoja1.Cells(nLin, 7) = psMarca
    xlHoja1.Range("C" & nLin & ":F" & nLin).Font.Bold = True
    nLin = nLin + 1
    
    xlHoja1.Cells(nLin, 1) = "OFICINA"
    xlHoja1.Cells(nLin, 3) = psAgencia
    xlHoja1.Cells(nLin, 5) = "AREA"
    xlHoja1.Cells(nLin, 6) = psArea
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & nLin & ":C" & nLin).Font.Bold = True
    nLin = nLin + 1
    
    xlHoja1.Cells(nLin, 1) = "ADQUISICION"
    xlHoja1.Cells(nLin, 3) = pdFechaAdq
    xlHoja1.Range("A" & nLin & ":B" & nLin).Merge True
    xlHoja1.Range("C" & 0 + nLin & ":C" & 0 + nLin).HorizontalAlignment = xlHAlignLeft

    nLin = nLin + 1
         
    xlHoja1.Cells(nLin, 1) = "FECHA"
    
    xlHoja1.Cells(nLin, 2) = "CONCEPTO"
    
    xlHoja1.Cells(nLin, 4) = "DEP MES"
      
    xlHoja1.Cells(nLin, 5) = "VALOR ACT"
    
    xlHoja1.Cells(nLin, 6) = "DEPRECIACION"
    
    xlHoja1.Cells(nLin, 7) = "VALOR"
    
    xlHoja1.Range("A" & nLin & ":G" & nLin).Font.Bold = True
    xlHoja1.Range("A" & nLin & ":G" & nLin).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A" & nLin & ":F" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("A" & nLin & ":F" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A" & nLin & ":F" & nLin).Borders(xlInsideVertical).Color = vbBlack
    'xlHoja1.Range("D6:E6").Borders(xlEdgeBottom).LineStyle = xlContinuous
    'xlHoja1.Range("Q6:S6").Borders(xlEdgeBottom).LineStyle = xlContinuous


    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""

        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
    End With
End Sub
'EJVG20130705 ***
Private Sub txtFechaIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFechaFin.SetFocus
    End If
End Sub
Private Sub txtFechaFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBS.SetFocus
    End If
End Sub
Private Sub txtBS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSerie.SetFocus
    End If
End Sub
Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdBuscar.SetFocus
    End If
End Sub
Private Sub txtSerie_EmiteDatos()
    lblSerieG.Caption = ""
    If txtSerie.Text <> "" Then
        lblSerieG.Caption = txtSerie.psDescripcion
    End If
End Sub
'END EJVG *******
