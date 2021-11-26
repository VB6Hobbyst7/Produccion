VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogReportes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   Icon            =   "frmLogReportes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdContab 
      Caption         =   "&Contabilidad"
      Height          =   300
      Left            =   5055
      TabIndex        =   3
      Top             =   1950
      Width           =   1185
   End
   Begin VB.CommandButton cmdAsiParcial 
      Caption         =   "&Asi. Parcial"
      Height          =   300
      Left            =   90
      TabIndex        =   15
      Top             =   1935
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdAsientosConsol 
      Caption         =   "&Asi. Consol"
      Height          =   300
      Left            =   1320
      TabIndex        =   14
      Top             =   1935
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdAsientos 
      Caption         =   "&Asientos"
      Height          =   300
      Left            =   2580
      TabIndex        =   13
      Top             =   1950
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MSComctlLib.ProgressBar Prg 
      Height          =   180
      Left            =   45
      TabIndex        =   4
      Top             =   2355
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   318
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   300
      Left            =   3825
      TabIndex        =   2
      Top             =   1950
      Width           =   1185
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   315
      Left            =   6300
      TabIndex        =   1
      Top             =   1950
      Width           =   1170
   End
   Begin VB.Frame fraFechas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   1845
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7455
      Begin Sicmact.TxtBuscar txtAlmacen 
         Height          =   315
         Left            =   210
         TabIndex        =   9
         Top             =   285
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Appearance      =   0
         BackColor       =   12648447
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
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   285
         Left            =   1425
         TabIndex        =   7
         Top             =   1485
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFecFin 
         Height          =   285
         Left            =   4605
         TabIndex        =   8
         Top             =   1500
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Sicmact.TxtBuscar txtPersona 
         Height          =   315
         Left            =   210
         TabIndex        =   16
         Top             =   930
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
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
         TipoBusqueda    =   3
      End
      Begin VB.Label lblProveedor 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   675
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblPersonaG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1635
         TabIndex        =   17
         Top             =   945
         Visible         =   0   'False
         Width           =   5715
      End
      Begin VB.Label lblAmacenG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1620
         TabIndex        =   10
         Top             =   300
         Width           =   5745
      End
      Begin VB.Label lblFecFin 
         Caption         =   "Fecha Final"
         Height          =   210
         Left            =   4605
         TabIndex        =   6
         Top             =   1260
         Width           =   1110
      End
      Begin VB.Label lblFecIni 
         Caption         =   "Fecha Inicial :  "
         Height          =   210
         Left            =   1455
         TabIndex        =   5
         Top             =   1260
         Width           =   1110
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flex 
      Height          =   375
      Left            =   2445
      TabIndex        =   11
      Top             =   1950
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   661
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   255
      Left            =   15
      SizeMode        =   1  'Stretch
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "frmLogReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim xlHoja2 As Excel.Worksheet

Dim lsCaption As String
Dim lbIngreso As Boolean
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
Dim lsPalabras As String
'*******************************

Public Sub Ini(pbIngreso As Boolean, psCaption As String)
    lbIngreso = pbIngreso
    lsCaption = psCaption
    Me.Show 1
End Sub

Private Sub cmdAsientos_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lsCadena As String
    Dim oAsiento As NContImprimir
    Set oAsiento = New NContImprimir

    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    
    Set rs = oALmacen.GetSalidasAlmacen(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), Me.txtAlmacen.Text)
        
    Prg.Max = rs.RecordCount + 2
    Prg.value = Prg.Min
        
    flex.Rows = 1
    flex.Cols = 8
    lsCadena = ""
    
    While Not rs.EOF
        lsCadena = lsCadena & oAsiento.ImprimeAsientoContable(rs!cMovNro, 60, 80) & oImpresora.gPrnSaltoPagina
        rs.MoveNext
        Prg.value = Prg.value + 1
    Wend
    
    oPrevio.Show lsCadena, Me.Caption, True
End Sub

Private Sub cmdAsientosConsol_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lsCadena As String
    Dim oAsiento As NContImprimir
    Set oAsiento = New NContImprimir
    
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    
    Set rs = oALmacen.GetSalidasAlmacen(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), Me.txtAlmacen.Text)
        
    Prg.Max = rs.RecordCount + 2
    Prg.value = Prg.Min
        
    flex.Rows = 1
    flex.Cols = 8
    lsCadena = ""
    
    While Not rs.EOF
        If lsCadena = "" Then
            lsCadena = Trim(Str(rs!nMovNro))
        Else
            lsCadena = lsCadena & "','" & Trim(Str(rs!nMovNro))
        End If
        rs.MoveNext
        Prg.value = Prg.value + 1
    Wend
    
    lsCadena = oAsiento.ImprimeAsientoConsolidado(lsCadena, gcMNDig, gdFecSis, "", "Asiento COnsolidado de Salidas de Logistica " & Me.mskFecIni & " - " & Me.mskFecFin.Text)
    
    oPrevio.Show lsCadena, Me.Caption, True
End Sub

Private Sub cmdAsiParcial_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lsCadena As String
    Dim oAsiento As NContImprimir
    Set oAsiento = New NContImprimir
    
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    
    Set rs = oALmacen.GetSalidasAlmacen(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), Me.txtAlmacen.Text)
        
    Prg.Max = rs.RecordCount + 2
    Prg.value = Prg.Min
        
    flex.Rows = 1
    flex.Cols = 8
    lsCadena = ""
    
    
    
    While Not rs.EOF
        If lsCadena = "" Then
            lsCadena = Trim(Str(rs!nMovNro))
        Else
            lsCadena = lsCadena & "','" & Trim(Str(rs!nMovNro))
        End If
        rs.MoveNext
        Prg.value = Prg.value + 1
    Wend
    
    lsCadena = oAsiento.ImprimeAsientoConsolidadoPalcial(lsCadena, gcMNDig, gdFecSis, "19", "Asiento COnsolidado de Salidas de Logistica " & Me.mskFecIni & " - " & Me.mskFecFin.Text)

    
    oPrevio.Show lsCadena, Me.Caption, True
End Sub

Private Sub cmdContab_Click()
    If lbIngreso Then
        IngresosContab
    Else
        SalidasContab
    End If
End Sub

Private Sub cmdProcesar_Click()
    If lbIngreso Then
        Ingresos
    Else
        Salidas
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.mskFecIni.SetFocus
End Sub

Private Sub Form_Load()
    Dim oDoc As DOperaciones
    Set oDoc = New DOperaciones
    
    Me.txtAlmacen.rs = oDoc.GetAlmacenes
    
    Me.txtAlmacen.Text = "1"
    Me.lblAmacenG.Caption = txtAlmacen.psDescripcion
    
    Caption = lsCaption
    
    If Not lbIngreso Then
        Me.cmdAsientos.Visible = True
        Me.cmdAsientosConsol.Visible = True
        Me.cmdAsiParcial.Visible = True
    Else
        Me.cmdAsientos.Visible = False
        Me.cmdAsientosConsol.Visible = False
        Me.cmdAsiParcial.Visible = False
        Me.lblProveedor.Visible = True
        Me.txtPersona.Visible = True
        Me.lblPersonaG.Visible = True
    End If
End Sub

Private Sub mskFecFin_GotFocus()
    mskFecFin.SelStart = 0
    mskFecFin.SelLength = 50
End Sub

Private Sub mskFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
    End If
End Sub

Private Sub mskFecIni_GotFocus()
    mskFecIni.SelStart = 0
    mskFecIni.SelLength = 50
End Sub

Private Sub mskFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskFecFin.SetFocus
    End If
End Sub

Private Sub txtAlmacen_EmiteDatos()
    Me.lblAmacenG.Caption = Me.txtAlmacen.psDescripcion
End Sub


'********************************
' Adiciona Hoja a LibroExcel
'********************************
Public Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet)
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = psHojName Then
       xlHoja1.Delete
       Exit For
    End If
Next
Set xlHoja1 = xlLibro.Worksheets.Add
xlHoja1.Name = psHojName
End Sub

'***********************************************************
' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
'***********************************************************
Private Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
        
Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application

If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
   Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  MsgBox Err.Description, vbInformation, "Aviso"
  ExcelBegin = False
End Function

'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Private Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub GeneraReporte()
    Dim I As Integer
    Dim K As Integer
    Dim J As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim sConec As String
    Dim lnAcum As Currency
    Dim VSQL As String
    
    Dim lnFilaMarcaIni As Integer
    Dim lnFilaMarcaFin As Integer
    
    Dim sTipoGara As String
    Dim sTipoCred As String
   
    lnFilaMarcaIni = 1
    
    xlHoja1.Columns.Range("A:A").Select
    xlHoja1.Columns.Range("A:A").NumberFormat = "@"
 
    For I = 0 To Me.flex.Rows - 1
        lnAcum = 0
        For J = 0 To Me.flex.Cols - 1
            xlHoja1.Cells(I + 1, J + 1) = Me.flex.TextMatrix(I, J)
            If I > 1 And J > 1 Then
                
                If IsNumeric(Me.flex.TextMatrix(I, J)) Then
                    lnAcum = lnAcum + CCur(Me.flex.TextMatrix(I, J))
                End If
            End If
        Next J
        
        If Me.flex.TextMatrix(I, 0) <> "" And I > 0 Then
            xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Select
        
            xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlDiagonalDown).LineStyle = xlNone
            xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlDiagonalUp).LineStyle = xlNone
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlInsideVertical).LineStyle = xlNone
            
            lnFilaMarcaIni = I + 1
        End If
        
        If I > 1 Then
            'xlHoja1.Range("A1:K" & Trim(Str(Me.flex.Rows))).Select
            'lnFilaMarcaIni = 0
            
            'VSQL = Format(lnAcum, "#,##0.00")  ' "=SUMA(" & Trim(ExcelColumnaString(3)) & Trim(I + 1) & ":" & Trim(ExcelColumnaString(Me.Flex.Cols)) & Trim(I + 1) & ")"
            'xlHoja1.Cells(I + 1, Me.Flex.Cols + 1).Formula = VSQL
            'xlHoja1.Cells(I + 1, Me.flex.Cols + 1) = VSQL
        End If
    Next I
        
    xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Select

    xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlDiagonalUp).LineStyle = xlNone
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlInsideVertical).LineStyle = xlNone
    
    lnFilaMarcaIni = I + 1
        
    xlHoja1.Range("A1:A" & Trim(Str(Me.flex.Rows))).Font.Bold = True
    xlHoja1.Range("B1:B" & Trim(Str(Me.flex.Rows))).Font.Bold = True
    xlHoja1.Range("G1:G" & Trim(Str(Me.flex.Rows))).Font.Bold = True
    xlHoja1.Range("1:1").Font.Bold = True

    xlHoja1.Range("E2:G" & Trim(Str(Me.flex.Rows))).NumberFormat = "#,##0.00"

    xlHoja1.Cells.Select
    xlHoja1.Cells.EntireColumn.AutoFit



'************************************

   With xlHoja1.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    xlHoja1.PageSetup.PrintArea = ""
    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = "&""Arial,Negrita""&18Listado de " & IIf(lbIngreso, "Ingresos", "Salidas") & " " & Format(CDate(Me.mskFecIni.Text), "mmmm yyyy")
        .RightHeader = "&P"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
'        .LeftMargin = Application.InchesToPoints(0)
'        .RightMargin = Application.InchesToPoints(0)
'        .TopMargin = Application.InchesToPoints(0.39)
'        .BottomMargin = Application.InchesToPoints(0)
'        .HeaderMargin = Application.InchesToPoints(0.14)
'        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 60
    End With
    
    
    xlHoja1.Columns("H:K").ColumnWidth = 45
    xlHoja1.Columns("C:C").ColumnWidth = 35
    xlHoja1.Columns("D:D").ColumnWidth = 45
    
    xlHoja1.Columns("H:K").Select
    With xlHoja1.Range("H:K")
        .HorizontalAlignment = xlJustify
        .VerticalAlignment = xlJustify
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    
End Sub

Private Sub Ingresos()
    Dim lnItem As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lsCadena As String
    Dim lnPagina As Long
    
    Dim lsFecha As String
    Dim lsDocumento As String
    Dim lsDescripcion As String
    Dim lsFactura As String
    Dim lsGuia As String
    Dim lsBienes As String
    Dim lsCantidad As String
    Dim lsImporte As String
    Dim lsPersona As String
    
    Dim lsDocAnt As String
    Dim lnAcumulador As Currency
    'Dim oPrevio As Previo.clsPrevio
    'Set oPrevio = New Previo.clsPrevio
    
    Dim lsArchivoN  As String
    Dim lbLibroOpen As Boolean
    
    
    If Not IsDate(Me.mskFecIni.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFecFin.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecFin.SetFocus
        Exit Sub
    End If
    
    Set rs = oALmacen.GetIngresosAlmacen(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), Trim(txtPersona.Text), Me.txtAlmacen.Text)
    
    If rs.RecordCount = 0 Then
       MsgBox "No existe información con respecto a la fecha Ingresada.", vbInformation, "Aviso"
       Exit Sub
    End If '***NAGL 20180423
    
    Prg.Max = rs.RecordCount + 2
    Prg.value = Prg.Min
    
    flex.Rows = 1
    flex.Cols = 11
    
    flex.TextMatrix(0, 0) = "Fecha"
    flex.TextMatrix(0, 1) = "Nro_Doc"
    flex.TextMatrix(0, 2) = "Persona"
    flex.TextMatrix(0, 3) = "Bienes"
    flex.TextMatrix(0, 4) = "Cantidad"
    flex.TextMatrix(0, 5) = "Importe"
    flex.TextMatrix(0, 6) = "Total"
    flex.TextMatrix(0, 7) = "Descripción"
    flex.TextMatrix(0, 8) = "Factura"
    flex.TextMatrix(0, 9) = "Guia"
    flex.TextMatrix(0, 10) = "CtaCont"
    
    lsCadena = ""
    lsCadena = lsCadena & CabeceraPagina("REPORTE DE INGRESO DE BIENES : " & Me.mskFecIni.Text & " - " & Me.mskFecFin, lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis)
    lsCadena = lsCadena & Encabezado("Fecha;10;Documento;15;Descripcion;50;Bienes;30;Cantidad;15;Importe;15;Persona;20;", lnItem, False)
    
    lsDocAnt = ""
    lnAcumulador = 0
    While Not rs.EOF
        flex.Rows = flex.Rows + 1
        If lsDocAnt <> rs!cDocNro Then
            If flex.Rows <> 2 Then
                lsImporte = Format(lnAcumulador, "#,##0.00")
                flex.TextMatrix(flex.Rows - 2, 6) = lsImporte
                lnAcumulador = 0
            End If
            
            lsFecha = Trim(Format(rs!dDocFecha, gsFormatoFechaView))
            lsDocumento = rs!cDocNro
            lsDescripcion = rs!cMovDesc
            lsFactura = rs!Factura & ""
            lsGuia = rs!GUIA & ""
            lsBienes = rs!cBSDescripcion
            lsCantidad = Format(rs!nMovCant, "#,##0.00")
            lsImporte = Format(rs!nMovImporte, "#,##0.00")
            lsPersona = rs!cPersNombre
        Else
            lsFecha = ""
            lsDocumento = ""
            lsDescripcion = ""
            lsBienes = rs!cBSDescripcion
            lsCantidad = Format(rs!nMovCant, "#,##0.00")
            lsImporte = Format(rs!nMovImporte, "#,##0.00")
            lsPersona = ""
            lsFactura = ""
            lsGuia = ""
        End If
        'lsCadena = lsCadena & Space(2) & lsFecha & Space(2) & lsDocumento & Space(2) & lsDescripcion & Space(2) & lsBienes & Space(2) & lsCantidad & Space(2) & lsImporte & Space(2) & lsPersona & oImpresora.gPrnSaltoLinea
        
        lnAcumulador = lnAcumulador + rs!nMovImporte
        
        flex.TextMatrix(flex.Rows - 1, 0) = lsFecha
        flex.TextMatrix(flex.Rows - 1, 1) = lsDocumento
        flex.TextMatrix(flex.Rows - 1, 2) = lsPersona
        flex.TextMatrix(flex.Rows - 1, 3) = lsBienes
        flex.TextMatrix(flex.Rows - 1, 4) = lsCantidad
        flex.TextMatrix(flex.Rows - 1, 5) = lsImporte
        flex.TextMatrix(flex.Rows - 1, 7) = lsDescripcion
        flex.TextMatrix(flex.Rows - 1, 8) = lsFactura
        flex.TextMatrix(flex.Rows - 1, 9) = lsGuia
        flex.TextMatrix(flex.Rows - 1, 10) = rs!cCtaContCod
        
        lsDocAnt = rs!cDocNro
        
        rs.MoveNext
        Prg.value = Prg.value + 1
    Wend
    
    If flex.Rows <> 2 Then
        lsImporte = Format(lnAcumulador, "#,##0.00")
        flex.TextMatrix(flex.Rows - 1, 6) = lsImporte
        lnAcumulador = 0
    End If
    'oPrevio.Show lsCadena, Caption, True
    
    Prg.value = Prg.Max
    
    'lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.mskFecFin.Text), "yyyymmdd") & ".xls"
     lsArchivoN = App.path & "\Spooler\" & gsCodUser & "_" & Format(CDate(Me.mskFecFin.Text), "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls" 'NAGL 20180424
    
    'OleExcel.Class = "ExcelWorkSheet"
    'lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    'If lbLibroOpen Then
       'Set xlHoja1 = xlLibro.Worksheets(1)
       'ExcelAddHoja Format(gdFecSis, "yyyymmdd") & "_" & gsCodUser, xlLibro, xlHoja1
       'Call GeneraReporte
       'OleExcel.Class = "ExcelWorkSheet"
       'ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       'OleExcel.SourceDoc = lsArchivoN
       'OleExcel.Verb = 1
       'OleExcel.Action = 1
       'OleExcel.DoVerb -1
    'End If
    'Comentado by NAGL 20180423
    
    'ARLO 20160126 ***
    If (gsopecod = 591504) Then
    lsPalabras = "Listado de Ingresos"
    ElseIf (gsopecod = 591502) Then
    lsPalabras = "Notas de Ingreso"
    ElseIf (gsopecod = 591503) Then
    lsPalabras = "Guias de Salida"
    End If
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Imprimio el Reporte " & lsPalabras & " del " & mskFecIni & " al " & mskFecFin
    Set objPista = Nothing
    '**************
    '*****NAGL 20180423*******
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro, False) 'NAGL Agregó False
    ExcelAddHoja Format(gdFecSis, "yyyymmdd") & "_" & gsCodUser, xlLibro, xlHoja1
    Call GeneraReporte
    ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
    Prg.value = Prg.Min
    If lsArchivoN <> "" Then
           CargaArchivo lsArchivoN, App.path & "\SPOOLER\"
    End If
    '****END NAGL 20180423*****
End Sub


Private Sub IngresosContab()
    Dim lnItem As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lsCadena As String
    Dim lnPagina As Long
    
    Dim lsFecha As String
    Dim lsDocumento As String
    Dim lsDescripcion As String
    Dim lsFactura As String
    Dim lsGuia As String
    Dim lsBienes As String
    Dim lsCantidad As String
    Dim lsImporte As String
    Dim lsPersona As String
    
    Dim lsDocAnt As String
    Dim lnAcumulador As Currency
    'Dim oPrevio As Previo.clsPrevio
    'Set oPrevio = New Previo.clsPrevio
    
    Dim lsArchivoN  As String
    Dim lbLibroOpen As Boolean
    
    
    If Not IsDate(Me.mskFecIni.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFecFin.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecFin.SetFocus
        Exit Sub
    End If
    
    'Set rs = oALmacen.GetIngresosAlmacen(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), Me.txtAlmacen.Text, True)
    '***********Agregado by NAGL 20180423
    Set rs = oALmacen.GetIngresosAlmacen(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), Trim(txtPersona.Text), Me.txtAlmacen.Text, True)
    If rs.RecordCount = 0 Then
       MsgBox "No existe información con respecto a la fecha Ingresada.", vbInformation, "Aviso"
       Exit Sub
    End If '***NAGL 20180423*************
            
    Prg.Max = rs.RecordCount + 2
    Prg.value = Prg.Min
        
    flex.Rows = 1
    flex.Cols = 11
    
    flex.TextMatrix(0, 0) = "Fecha"
    flex.TextMatrix(0, 1) = "Nro_Doc"
    flex.TextMatrix(0, 2) = "Persona"
    flex.TextMatrix(0, 3) = "Bienes"
    flex.TextMatrix(0, 4) = "Cantidad"
    flex.TextMatrix(0, 5) = "Importe"
    flex.TextMatrix(0, 6) = "Total"
    flex.TextMatrix(0, 7) = "Descripción"
    flex.TextMatrix(0, 8) = "Factura"
    flex.TextMatrix(0, 9) = "Guia"
    flex.TextMatrix(0, 10) = "CtaCont"
    
    lsCadena = ""
    lsCadena = lsCadena & CabeceraPagina("REPORTE DE INGRESO DE BIENES : " & Me.mskFecIni.Text & " - " & Me.mskFecFin, lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis)
    lsCadena = lsCadena & Encabezado("Fecha;10;Documento;15;Descripcion;50;Bienes;30;Cantidad;15;Importe;15;Persona;20;", lnItem, False)
    
    lsDocAnt = ""
    lnAcumulador = 0
    While Not rs.EOF
        flex.Rows = flex.Rows + 1
        If lsDocAnt <> rs!cCtaContCod Then
            If flex.Rows <> 2 Then
                lsImporte = Format(lnAcumulador, "#,##0.00")
                flex.TextMatrix(flex.Rows - 2, 6) = lsImporte
                lnAcumulador = 0
            End If
            
            lsFecha = Trim(Format(rs!dDocFecha, gsFormatoFechaView))
            lsDocumento = rs!cDocNro
            lsDescripcion = rs!cMovDesc
            lsFactura = rs!Factura & ""
            lsGuia = rs!GUIA & ""
            lsBienes = rs!cBSDescripcion
            lsCantidad = Format(rs!nMovCant, "#,##0.00")
            lsImporte = Format(rs!nMovImporte, "#,##0.00")
            lsPersona = rs!cPersNombre
        Else
            lsFecha = Trim(Format(rs!dDocFecha, gsFormatoFechaView))
            lsDocumento = rs!cDocNro
            lsDescripcion = rs!cMovDesc
            lsBienes = rs!cBSDescripcion
            lsCantidad = Format(rs!nMovCant, "#,##0.00")
            lsImporte = Format(rs!nMovImporte, "#,##0.00")
            lsPersona = rs!cPersNombre
            lsFactura = rs!Factura & ""
            lsGuia = rs!GUIA & ""
        End If
        'lsCadena = lsCadena & Space(2) & lsFecha & Space(2) & lsDocumento & Space(2) & lsDescripcion & Space(2) & lsBienes & Space(2) & lsCantidad & Space(2) & lsImporte & Space(2) & lsPersona & oImpresora.gPrnSaltoLinea
        
        lnAcumulador = lnAcumulador + rs!nMovImporte
        
        flex.TextMatrix(flex.Rows - 1, 0) = lsFecha
        flex.TextMatrix(flex.Rows - 1, 1) = lsDocumento
        flex.TextMatrix(flex.Rows - 1, 2) = lsPersona
        flex.TextMatrix(flex.Rows - 1, 3) = lsBienes
        flex.TextMatrix(flex.Rows - 1, 4) = lsCantidad
        flex.TextMatrix(flex.Rows - 1, 5) = lsImporte
        flex.TextMatrix(flex.Rows - 1, 7) = lsDescripcion
        flex.TextMatrix(flex.Rows - 1, 8) = lsFactura
        flex.TextMatrix(flex.Rows - 1, 9) = lsGuia
        flex.TextMatrix(flex.Rows - 1, 10) = rs!cCtaContCod
        
        lsDocAnt = rs!cCtaContCod
        
        rs.MoveNext
        Prg.value = Prg.value + 1
    Wend
    
    If flex.Rows <> 2 Then
        lsImporte = Format(lnAcumulador, "#,##0.00")
        flex.TextMatrix(flex.Rows - 1, 6) = lsImporte
        lnAcumulador = 0
    End If
    'oPrevio.Show lsCadena, Caption, True
    
    Prg.value = Prg.Max
    
    'lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.mskFecFin.Text), "yyyymmdd") & ".xls"
     lsArchivoN = App.path & "\Spooler\" & gsCodUser & "_" & Format(CDate(Me.mskFecFin.Text), "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls" 'NAGL 20180424
     
    'OleExcel.Class = "ExcelWorkSheet"
    'lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    'If lbLibroOpen Then
       'Set xlHoja1 = xlLibro.Worksheets(1)
       'ExcelAddHoja Format(gdFecSis, "yyyymmdd") & "_" & gsCodUser, xlLibro, xlHoja1
       'Call GeneraReporteContab
       'OleExcel.Class = "ExcelWorkSheet"
       'ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1

       'OleExcel.SourceDoc = lsArchivoN
       'OleExcel.Verb = 1
       'OleExcel.Action = 1
       'OleExcel.DoVerb -1
    'End If 'Comentado by NAGL 20180423
    
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    ExcelAddHoja Format(gdFecSis, "yyyymmdd") & "_" & gsCodUser, xlLibro, xlHoja1
    Call GeneraReporteContab
    ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
    Prg.value = Prg.Min
    If lsArchivoN <> "" Then
           CargaArchivo lsArchivoN, App.path & "\SPOOLER\"
    End If
    '****END NAGL 20180423*****
End Sub

Public Function ValidarFechas() As Boolean
     ValidarFechas = True
     If Not IsDate(mskFecIni) = True Then
        MsgBox "Ingrese Fecha Correcta", vbInformation, "Aviso"
        ValidarFechas = False
        Exit Function
     End If
     
     If Not IsDate(mskFecFin) = True Then
        MsgBox "Ingrese Fecha Correcta", vbInformation, "Aviso"
        ValidarFechas = False
        Exit Function
     End If

     If CDate(mskFecIni) > CDate(mskFecFin) Then
        MsgBox "La Fecha incial debe de ser Menor", vbInformation, "Aviso"
        ValidarFechas = False
        Exit Function
     End If
     
End Function


Private Sub Salidas()
    Dim lnItem As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lsCadena As String
    Dim lnPagina As Long
    
    Dim lsFecha As String
    Dim lsDocumento As String
    Dim lsDescripcion As String
    Dim lsBienes As String
    Dim lsCantidad As String
    Dim lsImporte As String
    Dim lsPersona As String
    
    Dim lsAgencia As String
    Dim lsArea As String
    Dim lsAACuenta As String
    
    Dim lsDocAnt As String
    Dim lnAcumulador As Currency
    'Dim oPrevio As Previo.clsPrevio
    'Set oPrevio = New Previo.clsPrevio
    
    Dim lsArchivoN  As String
    Dim lbLibroOpen As Boolean
    
 
    If ValidarFechas = False Then Exit Sub
    
    Set rs = oALmacen.GetSalidasAlmacen(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), Me.txtAlmacen.Text)
        
    Prg.Max = rs.RecordCount + 2
    Prg.value = Prg.Min
        
        
    flex.Rows = 1
    flex.Cols = 13
    
    flex.TextMatrix(0, 0) = "Fecha"
    
    flex.TextMatrix(0, 1) = "Agencia"
    flex.TextMatrix(0, 2) = "Area"
    flex.TextMatrix(0, 3) = "Codigo Gasto"
    
    flex.TextMatrix(0, 4) = "Nro_Doc"
    flex.TextMatrix(0, 5) = "Persona"
    
    flex.TextMatrix(0, 6) = "Codigo Bien"
    
    flex.TextMatrix(0, 7) = "Bienes"
    flex.TextMatrix(0, 8) = "Cantidad"
    flex.TextMatrix(0, 9) = "Importe"
    flex.TextMatrix(0, 10) = "Total"
    flex.TextMatrix(0, 11) = "Descripción"
    flex.TextMatrix(0, 12) = "CtaCont"
    
    lsCadena = ""
    lsCadena = lsCadena & CabeceraPagina("REPORTE DE INGRESO DE BIENES : " & Me.mskFecIni.Text & " - " & Me.mskFecFin, lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis)
    lsCadena = lsCadena & Encabezado("Fecha;10;Documento;15;Descripcion;50;Bienes;30;Cantidad;15;Importe;15;Persona;20;", lnItem, False)
    
    lsDocAnt = ""
    lnAcumulador = 0
    While Not rs.EOF
        flex.Rows = flex.Rows + 1
        If lsDocAnt <> rs!cDocNro Then
            If flex.Rows <> 2 Then
                lsImporte = Format(lnAcumulador, "#,##0.00")
                flex.TextMatrix(flex.Rows - 2, 9) = lsImporte
                lnAcumulador = 0
            End If
            
            lsFecha = Format(rs!dDocFecha, gsFormatoFechaView)
            lsDocumento = rs!cDocNro
            lsDescripcion = rs!cMovDesc
            lsBienes = rs!cBSDescripcion
            lsCantidad = Format(rs!nMovCant, "#,##0.00")
            lsImporte = Format(rs!nMovImporte, "#,##0.00")
            lsPersona = rs!cPersNombre & ""
            
            lsAgencia = rs!cAgeDescripcion
            lsArea = rs!cAreaDescripcion
            lsAACuenta = rs!cSubCtacod
        Else
            lsFecha = ""
            lsDocumento = ""
            lsDescripcion = ""
            lsBienes = rs!cBSDescripcion
            lsCantidad = Format(rs!nMovCant, "#,##0.00")
            lsImporte = Format(rs!nMovImporte, "#,##0.00")
            lsPersona = ""
            
            lsAgencia = rs!cAgeDescripcion
            lsArea = rs!cAreaDescripcion
            lsAACuenta = rs!cSubCtacod
        End If
        'lsCadena = lsCadena & Space(2) & lsFecha & Space(2) & lsDocumento & Space(2) & lsDescripcion & Space(2) & lsBienes & Space(2) & lsCantidad & Space(2) & lsImporte & Space(2) & lsPersona & oImpresora.gPrnSaltoLinea
        
        lnAcumulador = lnAcumulador + IIf(IsNull(rs!nMovImporte), 0, rs!nMovImporte)
        
        flex.TextMatrix(flex.Rows - 1, 0) = lsFecha
        
        flex.TextMatrix(flex.Rows - 1, 1) = lsAgencia
        flex.TextMatrix(flex.Rows - 1, 2) = lsArea
        flex.TextMatrix(flex.Rows - 1, 3) = lsAACuenta
        
        flex.TextMatrix(flex.Rows - 1, 4) = lsDocumento
        flex.TextMatrix(flex.Rows - 1, 5) = lsPersona
        
        flex.TextMatrix(flex.Rows - 1, 6) = rs!cBSCod
        
        flex.TextMatrix(flex.Rows - 1, 7) = lsBienes
        flex.TextMatrix(flex.Rows - 1, 8) = lsCantidad
        flex.TextMatrix(flex.Rows - 1, 9) = lsImporte
        flex.TextMatrix(flex.Rows - 1, 10) = lsImporte
        flex.TextMatrix(flex.Rows - 1, 11) = lsDescripcion
        flex.TextMatrix(flex.Rows - 1, 12) = rs!cCtaContCod
        
        
        lsDocAnt = rs!cDocNro
        
        rs.MoveNext
        Prg.value = Prg.value + 1
    Wend
    
    If flex.Rows <> 2 Then
        lsImporte = Format(lnAcumulador, "#,##0.00")
        flex.TextMatrix(flex.Rows - 1, 10) = lsImporte
        '**ALPA**31/03/2008
        '*************************************************************************************
        If flex.Rows - 2 >= 0 Then
            flex.TextMatrix(flex.Rows - 2, 10) = ""
        End If
        '**End********************************************************************************
        '*************************************************************************************
        'flex.TextMatrix(flex.Rows - 2, 10) = ""
        lnAcumulador = 0
    End If
    'oPrevio.Show lsCadena, Caption, True
    
    Prg.value = Prg.Max
    
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.mskFecFin.Text), "yyyymmdd") & ".xls"
    
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd") & "_" & gsCodUser, xlLibro, xlHoja1
       Call GeneraReporte
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1

       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
        'ARLO 20160126 ***
        If (gsopecod = 591505) Then
        lsPalabras = "Listado de Salidas"
        ElseIf (gsopecod = 591502) Then
        lsPalabras = "Notas de Ingreso"
        ElseIf (gsopecod = 591503) Then
        lsPalabras = "Guias de Salida"
        End If
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Imprimio el Reporte " & lsPalabras & " del " & mskFecIni & " al " & mskFecFin
        Set objPista = Nothing
        '**************
End Sub

Private Sub GeneraReporteContab()
    Dim I As Integer
    Dim K As Integer
    Dim J As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim sConec As String
    Dim lnAcum As Currency
    Dim VSQL As String
    Dim lsCtaContAnt As String
    Dim lnFilaMarcaIni As Integer
    Dim lnFilaMarcaFin As Integer
    
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnPosOpe As Integer
    
    lnFilaMarcaIni = 1
    
    If lbIngreso Then
        lnPosOpe = 10
    Else
        lnPosOpe = 8
    End If
    
    xlHoja1.Columns.Range("A:A").Select
    xlHoja1.Columns.Range("A:A").NumberFormat = "@"
 
    For I = 0 To Me.flex.Rows - 1
        lnAcum = 0
        For J = 0 To Me.flex.Cols - 1
            xlHoja1.Cells(I + 1, J + 1) = Me.flex.TextMatrix(I, J)
            If I > 1 And J > 1 Then
                
                If IsNumeric(Me.flex.TextMatrix(I, J)) Then
                    lnAcum = lnAcum + CCur(Me.flex.TextMatrix(I, J))
                End If
            End If
        Next J
        
        If Me.flex.TextMatrix(I, lnPosOpe) <> lsCtaContAnt And I > 0 Then
            xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Select
        
            xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlDiagonalDown).LineStyle = xlNone
            xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlDiagonalUp).LineStyle = xlNone
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(I + 0))).Borders(xlInsideVertical).LineStyle = xlNone
            
            lnFilaMarcaIni = I + 1
        End If
        
        lsCtaContAnt = Me.flex.TextMatrix(I, lnPosOpe)
    Next I
    
    
    xlHoja1.Range("A" & lnFilaMarcaIni & ":J" & Trim(Str(I + 0))).Select

    xlHoja1.Range("A" & lnFilaMarcaIni & ":J" & Trim(Str(I + 0))).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A" & lnFilaMarcaIni & ":J" & Trim(Str(I + 0))).Borders(xlDiagonalUp).LineStyle = xlNone
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":J" & Trim(Str(I + 0))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":J" & Trim(Str(I + 0))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":J" & Trim(Str(I + 0))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":J" & Trim(Str(I + 0))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    xlHoja1.Range("A" & lnFilaMarcaIni & ":J" & Trim(Str(I + 0))).Borders(xlInsideVertical).LineStyle = xlNone
    
    lnFilaMarcaIni = I + 1
        
    xlHoja1.Range("A1:A" & Trim(Str(Me.flex.Rows))).Font.Bold = True
    xlHoja1.Range("B1:B" & Trim(Str(Me.flex.Rows))).Font.Bold = True
    xlHoja1.Range("G1:G" & Trim(Str(Me.flex.Rows))).Font.Bold = True
    xlHoja1.Range("1:1").Font.Bold = True

    xlHoja1.Range("E2:G" & Trim(Str(Me.flex.Rows))).NumberFormat = "#,##0.00"

    xlHoja1.Cells.Select
    xlHoja1.Cells.EntireColumn.AutoFit



'************************************

   With xlHoja1.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    xlHoja1.PageSetup.PrintArea = ""
    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = "&""Arial,Negrita""&18Listado de " & IIf(lbIngreso, "Ingresos", "Salidas") & " " & Format(CDate(Me.mskFecIni.Text), "mmmm yyyy")
        .RightHeader = "&P"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
'        .LeftMargin = Application.InchesToPoints(0)
'        .RightMargin = Application.InchesToPoints(0)
'        .TopMargin = Application.InchesToPoints(0.39)
'        .BottomMargin = Application.InchesToPoints(0)
'        .HeaderMargin = Application.InchesToPoints(0.14)
'        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 60
    End With
    
    
    xlHoja1.Columns("H:H").ColumnWidth = 45
    xlHoja1.Columns("C:C").ColumnWidth = 35
    xlHoja1.Columns("D:D").ColumnWidth = 45
    
    xlHoja1.Columns("H:H").Select
    With xlHoja1.Range("H:H")
        .HorizontalAlignment = xlJustify
        .VerticalAlignment = xlJustify
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    
End Sub

Private Sub SalidasContab()
    Dim lnItem As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lsCadena As String
    Dim lnPagina As Long
    
    Dim lsFecha As String
    Dim lsDocumento As String
    Dim lsDescripcion As String
    Dim lsBienes As String
    Dim lsCantidad As String
    Dim lsImporte As String
    Dim lsPersona As String
    
    Dim lsDocAnt As String
    Dim lnAcumulador As Currency
    'Dim oPrevio As Previo.clsPrevio
    'Set oPrevio = New Previo.clsPrevio
    
    Dim lsArchivoN  As String
    Dim lbLibroOpen As Boolean
    
    Set rs = oALmacen.GetSalidasAlmacen(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), Me.txtAlmacen.Text, True)
        
    Prg.Max = rs.RecordCount + 2
    Prg.value = Prg.Min
        
    flex.Rows = 1
    flex.Cols = 9
    
    flex.TextMatrix(0, 0) = "Fecha"
    flex.TextMatrix(0, 1) = "Nro_Doc"
    flex.TextMatrix(0, 2) = "Persona"
    flex.TextMatrix(0, 3) = "Bienes"
    flex.TextMatrix(0, 4) = "Cantidad"
    flex.TextMatrix(0, 5) = "Importe"
    flex.TextMatrix(0, 6) = "Total"
    flex.TextMatrix(0, 7) = "Descripción"
    flex.TextMatrix(0, 8) = "CtaCont"
    
    lsCadena = ""
    lsCadena = lsCadena & CabeceraPagina("REPORTE DE INGRESO DE BIENES : " & Me.mskFecIni.Text & " - " & Me.mskFecFin, lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis)
    lsCadena = lsCadena & Encabezado("Fecha;10;Documento;15;Descripcion;50;Bienes;30;Cantidad;15;Importe;15;Persona;20;", lnItem, False)
    
    lsDocAnt = ""
    lnAcumulador = 0
    While Not rs.EOF
        flex.Rows = flex.Rows + 1
        If lsDocAnt <> rs!cCtaContCod Then
            If flex.Rows <> 2 Then
                lsImporte = Format(lnAcumulador, "#,##0.00")
                flex.TextMatrix(flex.Rows - 2, 6) = lsImporte
                lnAcumulador = 0
            End If
            
            lsFecha = Format(rs!dDocFecha, gsFormatoFechaView)
            lsDocumento = rs!cDocNro
            lsDescripcion = rs!cMovDesc
            lsBienes = rs!cBSDescripcion
            lsCantidad = Format(rs!nMovCant, "#,##0.00")
            lsImporte = Format(rs!nMovImporte, "#,##0.00")
            lsPersona = rs!cPersNombre & ""
        Else
            lsFecha = Format(rs!dDocFecha, gsFormatoFechaView)
            lsDocumento = rs!cDocNro
            lsDescripcion = rs!cMovDesc
            lsBienes = rs!cBSDescripcion
            lsCantidad = Format(rs!nMovCant, "#,##0.00")
            lsImporte = Format(rs!nMovImporte, "#,##0.00")
            lsPersona = rs!cPersNombre & ""
        End If
        'lsCadena = lsCadena & Space(2) & lsFecha & Space(2) & lsDocumento & Space(2) & lsDescripcion & Space(2) & lsBienes & Space(2) & lsCantidad & Space(2) & lsImporte & Space(2) & lsPersona & oImpresora.gPrnSaltoLinea
        
        lnAcumulador = lnAcumulador + IIf(IsNull(rs!nMovImporte), 0, rs!nMovImporte)
        
        flex.TextMatrix(flex.Rows - 1, 0) = lsFecha
        flex.TextMatrix(flex.Rows - 1, 1) = lsDocumento
        flex.TextMatrix(flex.Rows - 1, 2) = lsPersona
        flex.TextMatrix(flex.Rows - 1, 3) = lsBienes
        flex.TextMatrix(flex.Rows - 1, 4) = lsCantidad
        flex.TextMatrix(flex.Rows - 1, 5) = lsImporte
        flex.TextMatrix(flex.Rows - 1, 7) = lsDescripcion
        flex.TextMatrix(flex.Rows - 1, 8) = rs!cCtaContCod
        
        lsDocAnt = rs!cCtaContCod
        
        rs.MoveNext
        Prg.value = Prg.value + 1
    Wend
    
    If flex.Rows <> 2 Then
        lsImporte = Format(lnAcumulador, "#,##0.00")
        flex.TextMatrix(flex.Rows - 1, 6) = lsImporte
        lnAcumulador = 0
    End If
    'oPrevio.Show lsCadena, Caption, True
    
    Prg.value = Prg.Max
    
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.mskFecFin.Text), "yyyymmdd") & ".xls"
    
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd") & "_" & gsCodUser, xlLibro, xlHoja1
       Call GeneraReporteContab
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1

       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
End Sub

Private Sub txtPersona_EmiteDatos()
Me.lblPersonaG.Caption = Me.txtPersona.psDescripcion
End Sub

Private Sub txtPersona_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
    End If
End Sub
