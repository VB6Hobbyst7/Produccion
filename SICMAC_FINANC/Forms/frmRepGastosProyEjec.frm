VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmRepGastosProyEjec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Gastos Proyectodos vs Ejecutados"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "frmRepGastosProyEjec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame fraDatosReporte 
      Caption         =   "Datos "
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.ComboBox cboTpoRep 
         Height          =   315
         ItemData        =   "frmRepGastosProyEjec.frx":030A
         Left            =   3000
         List            =   "frmRepGastosProyEjec.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmRepGastosProyEjec.frx":0348
         Left            =   120
         List            =   "frmRepGastosProyEjec.frx":0373
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin Spinner.uSpinner txtAnio 
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Max             =   9999
         Min             =   1990
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblTipo 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblMes 
         Caption         =   "Mes:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblAnio 
         Alignment       =   2  'Center
         Caption         =   "Año:"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblGuion 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   255
      End
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblProc 
      Caption         =   "Procesando . . ."
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmRepGastosProyEjec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'***Nombre:         frmRepGastosProyEjec
'***Descripción:    Formulario que permite generar el reporte
'***                de gastos por proyectados y ejecutados.
'***Creación:       MIOL el 20130529 según ERS033-2013 OBJ C
'************************************************************
Option Explicit
'revisar
Option Base 0
Private Type TEstBal
    cCodCta  As String
    cDescrip As String
    cFormula As String
End Type
Private Type TCuentas
    cCta    As String
    nMES    As Double
    cDescrip As String
End Type
'probar
Private Type TMonto
    nMonto    As Double
    cGlosa As String
End Type

Dim EstBal() As TEstBal
Dim nContBal As Integer
'ALPA 20090512***************************
Dim EstBalReporte() As TEstBal
Dim nContBalReporte As Integer
Dim sCodOpeReporte As String
Dim CuentasReporte() As TCuentas
Dim MatrixReporte() As TCuentas
Dim nCuentasReporte As Integer
'****************************************
Dim Cuentas() As TCuentas
Dim MontOrd() As TMonto 'Probar
Dim nCuentas As Integer
Dim dFecha As Date
Dim sSql As String
Dim R As New ADODB.Recordset

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim sTipoRepoFormula As String
Dim sTituRepoFormula As String

Dim oNBal  As NBalanceCont
Dim oDBal  As DbalanceCont
Dim lnAnio As Integer
Dim lnAnioTC As Integer
Dim lnMes  As Integer
Dim lsRepCod  As String
Dim lsRepDesc As String

Dim fsCodReport As String
Dim lsFormaOrdenRep As String
Dim lbLimpiaDescrip As String

Dim nAnio As Integer

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGenerar_Click()
    If Me.cboMes.Text = "" Or Me.cboTpoRep.Text = "" Then
        MsgBox "Debe seleccionar los datos de mes y tipo de reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If Me.cboTpoRep.Text = "Comparativo Mes a Mes" Then
        GeneraReporteMes (1)
    ElseIf Me.cboTpoRep.Text = "Comparativo Consolidado" Then
        GeneraReporteMes (2)
    End If
End Sub

Private Sub Form_Load()
    lnMes = Mid(CStr(gdFecSis), 4, 2) - 1
    lnAnio = Right(CStr(gdFecSis), 4)
    Me.txtAnio.Valor = lnAnio
End Sub

Private Sub GeneraReporteMes(ByVal nTipoRep As Integer)
Dim oRepCtaColumna As DRepCtaColumna
Set oRepCtaColumna = New DRepCtaColumna
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim RSTEMP As ADODB.Recordset
Set RSTEMP = New ADODB.Recordset
Dim lsMoneda As String
Dim fs As Scripting.FileSystemObject
Dim xlAplicacion As Excel.Application
Dim lbExisteHoja As Boolean
Dim liLineas As Integer
Dim liLinDat As Integer
Dim liLinPro As Integer
Dim n As Integer
Dim nMES As Integer
Dim glsarchivo As String
Dim lsNomHoja As String
Dim obj_Excel As Object

    PB1.Min = 0
    PB1.Max = 7
    PB1.value = 0
    PB1.Visible = True
    Me.lblProc.Visible = True
    Set RSTEMP = oRepCtaColumna.GetRepProyectados(cboMes.ItemData(cboMes.ListIndex), txtAnio.Valor)

    If RSTEMP Is Nothing Then
        MsgBox "No exite informacion para imprimir", vbInformation, "Aviso"
        Exit Sub
    End If
    glsarchivo = "Reporte Proyectados vs Ejecutados" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsarchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsarchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets.Add
    PB1.value = 1
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    xlHoja1.PageSetup.Orientation = xlLandscape
            lbExisteHoja = False
            If nTipoRep = 1 Then
                lsNomHoja = "ProyEjecMes"
            ElseIf nTipoRep = 2 Then
                lsNomHoja = "ProyEjecCons"
            End If
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
    PB1.value = 2
            xlAplicacion.Range("A1:B10000").Font.Size = 9
            xlAplicacion.Range("A1:B10000").Font.Name = "Century Gothic"
            xlHoja1.Cells(1, 1) = "REPORTE DE GASTOS PROYECTADOS vs EJECUTADOS"
            If nTipoRep = 1 Then
                xlHoja1.Cells(2, 1) = "CMAC MAYNAS -  MENSUAL"
            ElseIf nTipoRep = 2 Then
                xlHoja1.Cells(2, 1) = "CMAC MAYNAS - CONSOLIDADO"
            End If
            If nTipoRep = 1 Then
                xlHoja1.Cells(1, 5) = "FECHA REPORTE"
                xlHoja1.Cells(2, 5) = gdFecSis
            End If
            xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 2)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 2)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 2)).Merge True
            xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 2)).Merge True
            xlHoja1.Range(xlHoja1.Cells(1, 5), xlHoja1.Cells(1, 5)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(2, 5), xlHoja1.Cells(2, 5)).Font.Bold = True
            liLineas = 4
            xlHoja1.Cells(liLineas, 1) = "NIVEL"
            xlHoja1.Cells(liLineas, 1).ColumnWidth = 5
            xlHoja1.Cells(liLineas, 1).Borders.LineStyle = 1
            xlHoja1.Cells(liLineas, 1).Font.Size = 8
            xlHoja1.Cells(liLineas, 2) = "DETALLE"
            xlHoja1.Cells(liLineas, 2).ColumnWidth = 40
            xlHoja1.Cells(liLineas, 2).Borders.LineStyle = 1
            xlHoja1.Cells(liLineas, 2).Font.Size = 8
            nMES = cboMes.ItemData(cboMes.ListIndex)
    PB1.value = 3
            If RSTEMP.RecordCount > 0 Then
                liLinDat = 5
                Do Until RSTEMP.EOF
                    xlHoja1.Cells(liLinDat, 1) = RSTEMP!nNivel
                    xlHoja1.Cells(liLinDat, 1).ColumnWidth = 5
                    xlHoja1.Cells(liLinDat, 1).Borders.LineStyle = 1
                    xlHoja1.Cells(liLinDat, 1).HorizontalAlignment = xlCenter
                    xlHoja1.Cells(liLinDat, 1).Font.Size = 8
                    Select Case RSTEMP!nNivel
                        Case 1: xlHoja1.Cells(liLinDat, 2) = RSTEMP!cConcepto
                                xlHoja1.Cells(liLinDat, 2).ColumnWidth = 40
                                xlHoja1.Cells(liLinDat, 2).Borders.LineStyle = 1
                                xlHoja1.Cells(liLinDat, 2).Font.Size = 8
                        Case 2: xlHoja1.Cells(liLinDat, 2) = "    " & RSTEMP!cConcepto
                                xlHoja1.Cells(liLinDat, 2).ColumnWidth = 40
                                xlHoja1.Cells(liLinDat, 2).Borders.LineStyle = 1
                                xlHoja1.Cells(liLinDat, 2).Font.Size = 8
                        Case 3: xlHoja1.Cells(liLinDat, 2) = "        " & RSTEMP!cConcepto
                                xlHoja1.Cells(liLinDat, 2).ColumnWidth = 40
                                xlHoja1.Cells(liLinDat, 2).Borders.LineStyle = 1
                                xlHoja1.Cells(liLinDat, 2).Font.Size = 8
                        Case 4: xlHoja1.Cells(liLinDat, 2) = "            " & RSTEMP!cConcepto
                                xlHoja1.Cells(liLinDat, 2).ColumnWidth = 40
                                xlHoja1.Cells(liLinDat, 2).Borders.LineStyle = 1
                                xlHoja1.Cells(liLinDat, 2).Font.Size = 8
                        Case 5: xlHoja1.Cells(liLinDat, 2) = "                    " & RSTEMP!cConcepto
                                xlHoja1.Cells(liLinDat, 2).ColumnWidth = 40
                                xlHoja1.Cells(liLinDat, 2).Borders.LineStyle = 1
                                xlHoja1.Cells(liLinDat, 2).Font.Size = 8
                        Case 6: xlHoja1.Cells(liLinDat, 2) = "                        " & RSTEMP!cConcepto
                                xlHoja1.Cells(liLinDat, 2).ColumnWidth = 40
                                xlHoja1.Cells(liLinDat, 2).Borders.LineStyle = 1
                                xlHoja1.Cells(liLinDat, 2).Font.Size = 8
                    End Select
                    liLinDat = liLinDat + 1
                    RSTEMP.MoveNext
                Loop
            End If
    PB1.value = 4
            CargaDatos (1)
            For n = 1 To nMES
                Select Case n
                            Case 1:
                                    liLinPro = 5
                                    xlHoja1.Cells(liLineas, 3) = "ENERO - PROY"
                                    xlHoja1.Cells(liLineas, 3).Borders.LineStyle = 1
                                    xlHoja1.Cells(liLineas, 3).ColumnWidth = 15
                                    xlHoja1.Cells(liLineas, 3).Font.Size = 8
                                    Set rs = oRepCtaColumna.GetRepProyectados(n, txtAnio.Valor)
                                    Do Until rs.EOF
                                        xlHoja1.Cells(liLinPro, 3) = rs!nMontoMes
                                        xlHoja1.Cells(liLinPro, 3).NumberFormat = "#,###0.00"
                                        xlHoja1.Cells(liLinPro, 3).Borders.LineStyle = 1
                                        xlHoja1.Cells(liLinPro, 3).ColumnWidth = 15
                                        xlHoja1.Cells(liLinPro, 3).Font.Size = 8
                                        liLinPro = liLinPro + 1
                                        rs.MoveNext
                                    Loop
                                    Set rs = Nothing
                                    GeneraReporteEjec n, 4, "ENERO - EJEC"
                            Case 2:
                                    liLinPro = 5
                                    xlHoja1.Cells(liLineas, 7) = "FEBRERO - PROY"
                                    xlHoja1.Cells(liLineas, 7).Borders.LineStyle = 1
                                    xlHoja1.Cells(liLineas, 7).ColumnWidth = 15
                                    xlHoja1.Cells(liLineas, 7).Font.Size = 8
                                    Set rs = oRepCtaColumna.GetRepProyectados(n, txtAnio.Valor)
                                    Do Until rs.EOF
                                        xlHoja1.Cells(liLinPro, 7) = rs!nMontoMes
                                        xlHoja1.Cells(liLinPro, 7).NumberFormat = "#,###0.00"
                                        xlHoja1.Cells(liLinPro, 7).Borders.LineStyle = 1
                                        xlHoja1.Cells(liLinPro, 7).ColumnWidth = 15
                                        xlHoja1.Cells(liLinPro, 7).Font.Size = 8
                                        liLinPro = liLinPro + 1
                                        rs.MoveNext
                                    Loop
                                    Set rs = Nothing
                                    GeneraReporteEjec n, 8, "FEBRERO - EJEC"
                            Case 3:
                                    liLinPro = 5
                                    xlHoja1.Cells(liLineas, 11) = "MARZO - PROY"
                                    xlHoja1.Cells(liLineas, 11).Borders.LineStyle = 1
                                    xlHoja1.Cells(liLineas, 11).ColumnWidth = 15
                                    xlHoja1.Cells(liLineas, 11).Font.Size = 8
                                    Set rs = oRepCtaColumna.GetRepProyectados(n, txtAnio.Valor)
                                    Do Until rs.EOF
                                        xlHoja1.Cells(liLinPro, 11) = rs!nMontoMes
                                        xlHoja1.Cells(liLinPro, 11).NumberFormat = "#,###0.00"
                                        xlHoja1.Cells(liLinPro, 11).Borders.LineStyle = 1
                                        xlHoja1.Cells(liLinPro, 11).ColumnWidth = 15
                                        xlHoja1.Cells(liLinPro, 11).Font.Size = 8
                                        liLinPro = liLinPro + 1
                                        rs.MoveNext
                                    Loop
                                    Set rs = Nothing
                                    GeneraReporteEjec n, 12, "MARZO - EJEC"
                            Case 4:
                                    liLinPro = 5
                                    xlHoja1.Cells(liLineas, 15) = "ABRIL - PROY"
                                    xlHoja1.Cells(liLineas, 15).Borders.LineStyle = 1
                                    xlHoja1.Cells(liLineas, 15).ColumnWidth = 15
                                    xlHoja1.Cells(liLineas, 15).Font.Size = 8
                                    Set rs = oRepCtaColumna.GetRepProyectados(n, txtAnio.Valor)
                                    Do Until rs.EOF
                                        xlHoja1.Cells(liLinPro, 15) = rs!nMontoMes
                                        xlHoja1.Cells(liLinPro, 15).NumberFormat = "#,###0.00"
                                        xlHoja1.Cells(liLinPro, 15).Borders.LineStyle = 1
                                        xlHoja1.Cells(liLinPro, 15).ColumnWidth = 15
                                        xlHoja1.Cells(liLinPro, 15).Font.Size = 8
                                        liLinPro = liLinPro + 1
                                        rs.MoveNext
                                    Loop
                                    Set rs = Nothing
                                    GeneraReporteEjec n, 16, "ABRIL - EJEC"
                            Case 5:
                                    liLinPro = 5
                                    xlHoja1.Cells(liLineas, 19) = "MAYO - PROY"
                                    xlHoja1.Cells(liLineas, 19).Borders.LineStyle = 1
                                    xlHoja1.Cells(liLineas, 19).ColumnWidth = 15
                                    xlHoja1.Cells(liLineas, 19).Font.Size = 8
                                    Set rs = oRepCtaColumna.GetRepProyectados(n, txtAnio.Valor)
                                    Do Until rs.EOF
                                        xlHoja1.Cells(liLinPro, 19) = rs!nMontoMes
                                        xlHoja1.Cells(liLinPro, 19).NumberFormat = "#,###0.00"
                                        xlHoja1.Cells(liLinPro, 19).Borders.LineStyle = 1
                                        xlHoja1.Cells(liLinPro, 19).ColumnWidth = 15
                                        xlHoja1.Cells(liLinPro, 19).Font.Size = 8
                                        liLinPro = liLinPro + 1
                                        rs.MoveNext
                                    Loop
                                    Set rs = Nothing
                                    GeneraReporteEjec n, 20, "MAYO - EJEC"
                            Case 6:
                                    liLinPro = 5
                                    xlHoja1.Cells(liLineas, 23) = "JUNIO - PROY"
                                    xlHoja1.Cells(liLineas, 23).Borders.LineStyle = 1
                                    xlHoja1.Cells(liLineas, 23).ColumnWidth = 15
                                    xlHoja1.Cells(liLineas, 23).Font.Size = 8
                                    Set rs = oRepCtaColumna.GetRepProyectados(n, txtAnio.Valor)
                                    Do Until rs.EOF
                                        xlHoja1.Cells(liLinPro, 23) = rs!nMontoMes
                                        xlHoja1.Cells(liLinPro, 23).NumberFormat = "#,###0.00"
                                        xlHoja1.Cells(liLinPro, 23).Borders.LineStyle = 1
                                        xlHoja1.Cells(liLinPro, 23).ColumnWidth = 15
                                        xlHoja1.Cells(liLinPro, 23).Font.Size = 8
                                        liLinPro = liLinPro + 1
                                        rs.MoveNext
                                    Loop
                                    Set rs = Nothing
                                    GeneraReporteEjec n, 24, "JUNIO - EJEC"
                            Case 7:
                                    liLinPro = 5
                                    xlHoja1.Cells(liLineas, 27) = "JULIO - PROY"
                                    xlHoja1.Cells(liLineas, 27).Borders.LineStyle = 1
                                    xlHoja1.Cells(liLineas, 27).ColumnWidth = 15
                                    xlHoja1.Cells(liLineas, 27).Font.Size = 8
                                    Set rs = oRepCtaColumna.GetRepProyectados(n, txtAnio.Valor)
                                    Do Until rs.EOF
                                        xlHoja1.Cells(liLinPro, 27) = rs!nMontoMes
                                        xlHoja1.Cells(liLinPro, 27).NumberFormat = "#,###0.00"
                                        xlHoja1.Cells(liLinPro, 27).Borders.LineStyle = 1
                                        xlHoja1.Cells(liLinPro, 27).ColumnWidth = 15
                                        xlHoja1.Cells(liLinPro, 27).Font.Size = 8
                                        liLinPro = liLinPro + 1
                                        rs.MoveNext
                                    Loop
                                    Set rs = Nothing
                                    GeneraReporteEjec n, 28, "JULIO - EJEC"
                            Case 8:
                                    liLinPro = 5
                                    xlHoja1.Cells(liLineas, 31) = "AGOSTO - PROY"
                                    xlHoja1.Cells(liLineas, 31).Borders.LineStyle = 1
                                    xlHoja1.Cells(liLineas, 31).ColumnWidth = 15
                                    xlHoja1.Cells(liLineas, 31).Font.Size = 8
                                    Set rs = oRepCtaColumna.GetRepProyectados(n, txtAnio.Valor)
                                    Do Until rs.EOF
                                        xlHoja1.Cells(liLinPro, 31) = rs!nMontoMes
                                        xlHoja1.Cells(liLinPro, 31).NumberFormat = "#,###0.00"
                                        xlHoja1.Cells(liLinPro, 31).Borders.LineStyle = 1
                                        xlHoja1.Cells(liLinPro, 31).ColumnWidth = 15
                                        xlHoja1.Cells(liLinPro, 31).Font.Size = 8
                                        liLinPro = liLinPro + 1
                                        rs.MoveNext
                                    Loop
                                    Set rs = Nothing
                                    GeneraReporteEjec n, 32, "AGOSTO - EJEC"
                            Case 9:
                                    liLinPro = 5
                                    xlHoja1.Cells(liLineas, 35) = "SETIEMBRE - PROY"
                                    xlHoja1.Cells(liLineas, 35).Borders.LineStyle = 1
                                    xlHoja1.Cells(liLineas, 35).ColumnWidth = 15
                                    xlHoja1.Cells(liLineas, 35).Font.Size = 8
                                    Set rs = oRepCtaColumna.GetRepProyectados(n, txtAnio.Valor)
                                    Do Until rs.EOF
                                        xlHoja1.Cells(liLinPro, 35) = rs!nMontoMes
                                        xlHoja1.Cells(liLinPro, 35).NumberFormat = "#,###0.00"
                                        xlHoja1.Cells(liLinPro, 35).Borders.LineStyle = 1
                                        xlHoja1.Cells(liLinPro, 35).ColumnWidth = 15
                                        xlHoja1.Cells(liLinPro, 35).Font.Size = 8
                                        liLinPro = liLinPro + 1
                                        rs.MoveNext
                                    Loop
                                    Set rs = Nothing
                                    GeneraReporteEjec n, 36, "SETIEMBRE - EJEC"
                            Case 10:
                                    liLinPro = 5
                                    xlHoja1.Cells(liLineas, 39) = "OCTUBRE - PROY"
                                    xlHoja1.Cells(liLineas, 39).Borders.LineStyle = 1
                                    xlHoja1.Cells(liLineas, 39).ColumnWidth = 15
                                    xlHoja1.Cells(liLineas, 39).Font.Size = 8
                                    Set rs = oRepCtaColumna.GetRepProyectados(n, txtAnio.Valor)
                                    Do Until rs.EOF
                                        xlHoja1.Cells(liLinPro, 39) = rs!nMontoMes
                                        xlHoja1.Cells(liLinPro, 39).NumberFormat = "#,###0.00"
                                        xlHoja1.Cells(liLinPro, 39).Borders.LineStyle = 1
                                        xlHoja1.Cells(liLinPro, 39).ColumnWidth = 15
                                        xlHoja1.Cells(liLinPro, 39).Font.Size = 8
                                        liLinPro = liLinPro + 1
                                        rs.MoveNext
                                    Loop
                                    Set rs = Nothing
                                    GeneraReporteEjec n, 40, "OCTUBRE - EJEC"
                            Case 11:
                                    liLinPro = 5
                                    xlHoja1.Cells(liLineas, 43) = "NOVIEMBRE - PROY"
                                    xlHoja1.Cells(liLineas, 43).Borders.LineStyle = 1
                                    xlHoja1.Cells(liLineas, 43).ColumnWidth = 15
                                    xlHoja1.Cells(liLineas, 43).Font.Size = 8
                                    Set rs = oRepCtaColumna.GetRepProyectados(n, txtAnio.Valor)
                                    Do Until rs.EOF
                                        xlHoja1.Cells(liLinPro, 43) = rs!nMontoMes
                                        xlHoja1.Cells(liLinPro, 43).NumberFormat = "#,###0.00"
                                        xlHoja1.Cells(liLinPro, 43).Borders.LineStyle = 1
                                        xlHoja1.Cells(liLinPro, 43).ColumnWidth = 15
                                        xlHoja1.Cells(liLinPro, 43).Font.Size = 8
                                        liLinPro = liLinPro + 1
                                        rs.MoveNext
                                    Loop
                                    Set rs = Nothing
                                    GeneraReporteEjec n, 44, "NOVIEMBRE - EJEC"
                            Case 12:
                                    liLinPro = 5
                                    xlHoja1.Cells(liLineas, 47) = "DICIEMBRE - PROY"
                                    xlHoja1.Cells(liLineas, 47).Borders.LineStyle = 1
                                    xlHoja1.Cells(liLineas, 47).ColumnWidth = 15
                                    xlHoja1.Cells(liLineas, 47).Font.Size = 8
                                    Set rs = oRepCtaColumna.GetRepProyectados(n, txtAnio.Valor)
                                    Do Until rs.EOF
                                        xlHoja1.Cells(liLinPro, 47) = rs!nMontoMes
                                        xlHoja1.Cells(liLinPro, 47).NumberFormat = "#,###0.00"
                                        xlHoja1.Cells(liLinPro, 47).Borders.LineStyle = 1
                                        xlHoja1.Cells(liLinPro, 47).ColumnWidth = 15
                                        xlHoja1.Cells(liLinPro, 47).Font.Size = 8
                                        
                                        liLinPro = liLinPro + 1
                                        rs.MoveNext
                                    Loop
                                    Set rs = Nothing
                                    GeneraReporteEjec n, 48, "DICIEMBRE - EJEC"
                End Select
            Next
    PB1.value = 5
            Set rs = oRepCtaColumna.GetDatosProyEjec(nMES, txtAnio.Valor)
                Dim nColCom As Integer
                Dim nCantFila As Integer
                nCantFila = rs.RecordCount
                Select Case nMES
                        Case 1: nColCom = 7 'ENERO
                            liLinPro = 5 'PROYECTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom) = "ENERO PROY. CONSOL"
                                FormatoRep liLinPro, nColCom, 0
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom) = xlHoja1.Cells(liLinPro, 3)
                                FormatoRep liLinPro, nColCom, 4
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'EJECUTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom + 1) = "ENERO EJEC. CONSOL"
                                FormatoRep liLinPro, nColCom, 1
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 1) = xlHoja1.Cells(liLinPro, 4)
                                FormatoRep liLinPro, nColCom, 5
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'DIFERENCIA
                                xlHoja1.Cells(liLinPro - 1, nColCom + 2) = "DIFERENCIA"
                                FormatoRep liLinPro, nColCom, 2
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 2) = xlHoja1.Cells(liLinPro, 5)
                                FormatoRep liLinPro, nColCom, 6
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'PORCENTAJE
                                xlHoja1.Cells(liLinPro - 1, nColCom + 3) = "PORCENTAJE"
                                FormatoRep liLinPro, nColCom, 3
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 3) = xlHoja1.Cells(liLinPro, 6)
                                FormatoRep liLinPro, nColCom, 7
                                liLinPro = liLinPro + 1
                            Next
                        Case 2: nColCom = 11 'FEBRERO
                            liLinPro = 5 'PROYECTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom) = "FEBRERO PROY. CONSOL"
                                FormatoRep liLinPro, nColCom, 0
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom) = xlHoja1.Cells(liLinPro, 3) + xlHoja1.Cells(liLinPro, 7)
                                FormatoRep liLinPro, nColCom, 4
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'EJECUTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom + 1) = "FEBRERO EJEC. CONSOL"
                                FormatoRep liLinPro, nColCom, 1
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 1) = xlHoja1.Cells(liLinPro, 4) + xlHoja1.Cells(liLinPro, 8)
                                FormatoRep liLinPro, nColCom, 5
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'DIFERENCIA
                                xlHoja1.Cells(liLinPro - 1, nColCom + 2) = "DIFERENCIA"
                                FormatoRep liLinPro, nColCom, 2
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 2) = xlHoja1.Cells(liLinPro, 5) + xlHoja1.Cells(liLinPro, 9)
                                FormatoRep liLinPro, nColCom, 6
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'PORCENTAJE
                                xlHoja1.Cells(liLinPro - 1, nColCom + 3) = "PORCENTAJE"
                                FormatoRep liLinPro, nColCom, 3
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 3) = (xlHoja1.Cells(liLinPro, 6) + xlHoja1.Cells(liLinPro, 10)) / 2
                                FormatoRep liLinPro, nColCom, 7
                                liLinPro = liLinPro + 1
                            Next
                        Case 3: nColCom = 15 'MARZO
                            liLinPro = 5 'PROYECTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom) = "MARZO PROY. CONSOL"
                                FormatoRep liLinPro, nColCom, 0
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom) = xlHoja1.Cells(liLinPro, 3) + xlHoja1.Cells(liLinPro, 7) + xlHoja1.Cells(liLinPro, 11)
                                FormatoRep liLinPro, nColCom, 4
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'EJECUTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom + 1) = "MARZO EJEC. CONSOL"
                                FormatoRep liLinPro, nColCom, 1
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 1) = xlHoja1.Cells(liLinPro, 4) + xlHoja1.Cells(liLinPro, 8) + xlHoja1.Cells(liLinPro, 12)
                                FormatoRep liLinPro, nColCom, 5
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'DIFERENCIA
                                xlHoja1.Cells(liLinPro - 1, nColCom + 2) = "DIFERENCIA"
                                FormatoRep liLinPro, nColCom, 2
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 2) = xlHoja1.Cells(liLinPro, 5) + xlHoja1.Cells(liLinPro, 9) + xlHoja1.Cells(liLinPro, 13)
                                FormatoRep liLinPro, nColCom, 6
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'PORCENTAJE
                                xlHoja1.Cells(liLinPro - 1, nColCom + 3) = "PORCENTAJE"
                                FormatoRep liLinPro, nColCom, 3
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 3) = (xlHoja1.Cells(liLinPro, 6) + xlHoja1.Cells(liLinPro, 10) + xlHoja1.Cells(liLinPro, 14)) / 3
                                FormatoRep liLinPro, nColCom, 7
                                liLinPro = liLinPro + 1
                            Next
                        Case 4: nColCom = 19 'ABRIL
                            liLinPro = 5 'PROYECTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom) = "ABRIL PROY. CONSOL"
                                FormatoRep liLinPro, nColCom, 0
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom) = xlHoja1.Cells(liLinPro, 3) + xlHoja1.Cells(liLinPro, 7) + xlHoja1.Cells(liLinPro, 11) + xlHoja1.Cells(liLinPro, 15)
                                FormatoRep liLinPro, nColCom, 4
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'EJECUTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom + 1) = "ABRIL EJEC. CONSOL"
                                FormatoRep liLinPro, nColCom, 1
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 1) = xlHoja1.Cells(liLinPro, 4) + xlHoja1.Cells(liLinPro, 8) + xlHoja1.Cells(liLinPro, 12) + xlHoja1.Cells(liLinPro, 16)
                                FormatoRep liLinPro, nColCom, 5
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'DIFERENCIA
                                xlHoja1.Cells(liLinPro - 1, nColCom + 2) = "DIFERENCIA"
                                FormatoRep liLinPro, nColCom, 2
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 2) = xlHoja1.Cells(liLinPro, 5) + xlHoja1.Cells(liLinPro, 9) + xlHoja1.Cells(liLinPro, 13) + xlHoja1.Cells(liLinPro, 17)
                                FormatoRep liLinPro, nColCom, 6
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'PORCENTAJE
                                xlHoja1.Cells(liLinPro - 1, nColCom + 3) = "PORCENTAJE"
                                FormatoRep liLinPro, nColCom, 3
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 3) = (xlHoja1.Cells(liLinPro, 6) + xlHoja1.Cells(liLinPro, 10) + xlHoja1.Cells(liLinPro, 14) + xlHoja1.Cells(liLinPro, 18)) / 4
                                FormatoRep liLinPro, nColCom, 7
                                liLinPro = liLinPro + 1
                            Next
                        Case 5: nColCom = 23 'MAYO
                            liLinPro = 5 'PROYECTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom) = "MAYO PROY. CONSOL"
                                FormatoRep liLinPro, nColCom, 0
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom) = xlHoja1.Cells(liLinPro, 3) + xlHoja1.Cells(liLinPro, 7) + xlHoja1.Cells(liLinPro, 11) + xlHoja1.Cells(liLinPro, 15) + xlHoja1.Cells(liLinPro, 19)
                                FormatoRep liLinPro, nColCom, 4
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'EJECUTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom + 1) = "MAYO EJEC. CONSOL"
                                FormatoRep liLinPro, nColCom, 1
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 1) = xlHoja1.Cells(liLinPro, 4) + xlHoja1.Cells(liLinPro, 8) + xlHoja1.Cells(liLinPro, 12) + xlHoja1.Cells(liLinPro, 16) + xlHoja1.Cells(liLinPro, 20)
                                FormatoRep liLinPro, nColCom, 5
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'DIFERENCIA
                                xlHoja1.Cells(liLinPro - 1, nColCom + 2) = "DIFERENCIA"
                                FormatoRep liLinPro, nColCom, 2
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 2) = xlHoja1.Cells(liLinPro, 5) + xlHoja1.Cells(liLinPro, 9) + xlHoja1.Cells(liLinPro, 13) + xlHoja1.Cells(liLinPro, 17) + xlHoja1.Cells(liLinPro, 21)
                                FormatoRep liLinPro, nColCom, 6
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'PORCENTAJE
                                xlHoja1.Cells(liLinPro - 1, nColCom + 3) = "PORCENTAJE"
                                FormatoRep liLinPro, nColCom, 3
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 3) = (xlHoja1.Cells(liLinPro, 6) + xlHoja1.Cells(liLinPro, 10) + xlHoja1.Cells(liLinPro, 14) + xlHoja1.Cells(liLinPro, 18) + xlHoja1.Cells(liLinPro, 22)) / 5
                                FormatoRep liLinPro, nColCom, 7
                                liLinPro = liLinPro + 1
                            Next
                        Case 6: nColCom = 27 'JUNIO
                            liLinPro = 5 'PROYECTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom) = "JUNIO PROY. CONSOL"
                                FormatoRep liLinPro, nColCom, 0
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom) = xlHoja1.Cells(liLinPro, 3) + xlHoja1.Cells(liLinPro, 7) + xlHoja1.Cells(liLinPro, 11) + xlHoja1.Cells(liLinPro, 15) + xlHoja1.Cells(liLinPro, 19) + xlHoja1.Cells(liLinPro, 23)
                                FormatoRep liLinPro, nColCom, 4
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'EJECUTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom + 1) = "JUNIO EJEC. CONSOL"
                                FormatoRep liLinPro, nColCom, 1
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 1) = xlHoja1.Cells(liLinPro, 4) + xlHoja1.Cells(liLinPro, 8) + xlHoja1.Cells(liLinPro, 12) + xlHoja1.Cells(liLinPro, 16) + xlHoja1.Cells(liLinPro, 20) + xlHoja1.Cells(liLinPro, 24)
                                FormatoRep liLinPro, nColCom, 5
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'DIFERENCIA
                                xlHoja1.Cells(liLinPro - 1, nColCom + 2) = "DIFERENCIA"
                                FormatoRep liLinPro, nColCom, 2
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 2) = xlHoja1.Cells(liLinPro, 5) + xlHoja1.Cells(liLinPro, 9) + xlHoja1.Cells(liLinPro, 13) + xlHoja1.Cells(liLinPro, 17) + xlHoja1.Cells(liLinPro, 21) + xlHoja1.Cells(liLinPro, 25)
                                FormatoRep liLinPro, nColCom, 6
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'PORCENTAJE
                                xlHoja1.Cells(liLinPro - 1, nColCom + 3) = "PORCENTAJE"
                                FormatoRep liLinPro, nColCom, 3
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 3) = (xlHoja1.Cells(liLinPro, 6) + xlHoja1.Cells(liLinPro, 10) + xlHoja1.Cells(liLinPro, 14) + xlHoja1.Cells(liLinPro, 18) + xlHoja1.Cells(liLinPro, 22) + xlHoja1.Cells(liLinPro, 26)) / 6
                                FormatoRep liLinPro, nColCom, 7
                                liLinPro = liLinPro + 1
                            Next
                        Case 7: nColCom = 31 'JULIO
                            liLinPro = 5 'PROYECTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom) = "JULIO PROY. CONSOL"
                                FormatoRep liLinPro, nColCom, 0
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom) = xlHoja1.Cells(liLinPro, 3) + xlHoja1.Cells(liLinPro, 7) + xlHoja1.Cells(liLinPro, 11) + xlHoja1.Cells(liLinPro, 15) + xlHoja1.Cells(liLinPro, 19) + xlHoja1.Cells(liLinPro, 23) + xlHoja1.Cells(liLinPro, 27)
                                FormatoRep liLinPro, nColCom, 4
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'EJECUTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom + 1) = "JULIO EJEC. CONSOL"
                                FormatoRep liLinPro, nColCom, 1
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 1) = xlHoja1.Cells(liLinPro, 4) + xlHoja1.Cells(liLinPro, 8) + xlHoja1.Cells(liLinPro, 12) + xlHoja1.Cells(liLinPro, 16) + xlHoja1.Cells(liLinPro, 20) + xlHoja1.Cells(liLinPro, 24) + xlHoja1.Cells(liLinPro, 28)
                                FormatoRep liLinPro, nColCom, 5
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'DIFERENCIA
                                xlHoja1.Cells(liLinPro - 1, nColCom + 2) = "DIFERENCIA"
                                FormatoRep liLinPro, nColCom, 2
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 2) = xlHoja1.Cells(liLinPro, 5) + xlHoja1.Cells(liLinPro, 9) + xlHoja1.Cells(liLinPro, 13) + xlHoja1.Cells(liLinPro, 17) + xlHoja1.Cells(liLinPro, 21) + xlHoja1.Cells(liLinPro, 25) + xlHoja1.Cells(liLinPro, 29)
                                FormatoRep liLinPro, nColCom, 6
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'PORCENTAJE
                                xlHoja1.Cells(liLinPro - 1, nColCom + 3) = "PORCENTAJE"
                                FormatoRep liLinPro, nColCom, 3
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 3) = (xlHoja1.Cells(liLinPro, 6) + xlHoja1.Cells(liLinPro, 10) + xlHoja1.Cells(liLinPro, 14) + xlHoja1.Cells(liLinPro, 18) + xlHoja1.Cells(liLinPro, 22) + xlHoja1.Cells(liLinPro, 26) + xlHoja1.Cells(liLinPro, 30)) / 7
                                FormatoRep liLinPro, nColCom, 7
                                liLinPro = liLinPro + 1
                            Next
                        Case 8: nColCom = 35 'AGOSTO
                            liLinPro = 5 'PROYECTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom) = "AGOSTO PROY. CONSOL"
                                FormatoRep liLinPro, nColCom, 0
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom) = xlHoja1.Cells(liLinPro, 3) + xlHoja1.Cells(liLinPro, 7) + xlHoja1.Cells(liLinPro, 11) + xlHoja1.Cells(liLinPro, 15) + xlHoja1.Cells(liLinPro, 19) + xlHoja1.Cells(liLinPro, 23) + xlHoja1.Cells(liLinPro, 27) + xlHoja1.Cells(liLinPro, 31)
                                FormatoRep liLinPro, nColCom, 4
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'EJECUTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom + 1) = "AGOSTO EJEC. CONSOL"
                                FormatoRep liLinPro, nColCom, 1
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 1) = xlHoja1.Cells(liLinPro, 4) + xlHoja1.Cells(liLinPro, 8) + xlHoja1.Cells(liLinPro, 12) + xlHoja1.Cells(liLinPro, 16) + xlHoja1.Cells(liLinPro, 20) + xlHoja1.Cells(liLinPro, 24) + xlHoja1.Cells(liLinPro, 28) + xlHoja1.Cells(liLinPro, 32)
                                FormatoRep liLinPro, nColCom, 5
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'DIFERENCIA
                                xlHoja1.Cells(liLinPro - 1, nColCom + 2) = "DIFERENCIA"
                                FormatoRep liLinPro, nColCom, 2
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 2) = xlHoja1.Cells(liLinPro, 5) + xlHoja1.Cells(liLinPro, 9) + xlHoja1.Cells(liLinPro, 13) + xlHoja1.Cells(liLinPro, 17) + xlHoja1.Cells(liLinPro, 21) + xlHoja1.Cells(liLinPro, 25) + xlHoja1.Cells(liLinPro, 29) + xlHoja1.Cells(liLinPro, 33)
                                FormatoRep liLinPro, nColCom, 6
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'PORCENTAJE
                                xlHoja1.Cells(liLinPro - 1, nColCom + 3) = "PORCENTAJE"
                                FormatoRep liLinPro, nColCom, 3
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 3) = (xlHoja1.Cells(liLinPro, 6) + xlHoja1.Cells(liLinPro, 10) + xlHoja1.Cells(liLinPro, 14) + xlHoja1.Cells(liLinPro, 18) + xlHoja1.Cells(liLinPro, 22) + xlHoja1.Cells(liLinPro, 26) + xlHoja1.Cells(liLinPro, 30) + xlHoja1.Cells(liLinPro, 34)) / 8
                                FormatoRep liLinPro, nColCom, 7
                                liLinPro = liLinPro + 1
                            Next
                        Case 9: nColCom = 39 'SETIEMBRE
                            liLinPro = 5 'PROYECTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom) = "SETIEMBRE PROY. CONSOL"
                                FormatoRep liLinPro, nColCom, 0
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom) = xlHoja1.Cells(liLinPro, 3) + xlHoja1.Cells(liLinPro, 7) + xlHoja1.Cells(liLinPro, 11) + xlHoja1.Cells(liLinPro, 15) + xlHoja1.Cells(liLinPro, 19) + xlHoja1.Cells(liLinPro, 23) + xlHoja1.Cells(liLinPro, 27) + xlHoja1.Cells(liLinPro, 31) + xlHoja1.Cells(liLinPro, 35)
                                FormatoRep liLinPro, nColCom, 4
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'EJECUTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom + 1) = "SETIEMBRE EJEC. CONSOL"
                                FormatoRep liLinPro, nColCom, 1
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 1) = xlHoja1.Cells(liLinPro, 4) + xlHoja1.Cells(liLinPro, 8) + xlHoja1.Cells(liLinPro, 12) + xlHoja1.Cells(liLinPro, 16) + xlHoja1.Cells(liLinPro, 20) + xlHoja1.Cells(liLinPro, 24) + xlHoja1.Cells(liLinPro, 28) + xlHoja1.Cells(liLinPro, 32) + xlHoja1.Cells(liLinPro, 36)
                                FormatoRep liLinPro, nColCom, 5
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'DIFERENCIA
                                xlHoja1.Cells(liLinPro - 1, nColCom + 2) = "DIFERENCIA"
                                FormatoRep liLinPro, nColCom, 2
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 2) = xlHoja1.Cells(liLinPro, 5) + xlHoja1.Cells(liLinPro, 9) + xlHoja1.Cells(liLinPro, 13) + xlHoja1.Cells(liLinPro, 17) + xlHoja1.Cells(liLinPro, 21) + xlHoja1.Cells(liLinPro, 25) + xlHoja1.Cells(liLinPro, 29) + xlHoja1.Cells(liLinPro, 33) + xlHoja1.Cells(liLinPro, 37)
                                FormatoRep liLinPro, nColCom, 6
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'PORCENTAJE
                                xlHoja1.Cells(liLinPro - 1, nColCom + 3) = "PORCENTAJE"
                                FormatoRep liLinPro, nColCom, 3
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 3) = (xlHoja1.Cells(liLinPro, 6) + xlHoja1.Cells(liLinPro, 10) + xlHoja1.Cells(liLinPro, 14) + xlHoja1.Cells(liLinPro, 18) + xlHoja1.Cells(liLinPro, 22) + xlHoja1.Cells(liLinPro, 26) + xlHoja1.Cells(liLinPro, 30) + xlHoja1.Cells(liLinPro, 34) + xlHoja1.Cells(liLinPro, 38)) / 9
                                FormatoRep liLinPro, nColCom, 7
                                liLinPro = liLinPro + 1
                            Next
                        Case 10: nColCom = 43 'OCTUBRE
                            liLinPro = 5 'PROYECTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom) = "OCTUBRE PROY. CONSOL"
                                FormatoRep liLinPro, nColCom, 0
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom) = xlHoja1.Cells(liLinPro, 3) + xlHoja1.Cells(liLinPro, 7) + xlHoja1.Cells(liLinPro, 11) + xlHoja1.Cells(liLinPro, 15) + xlHoja1.Cells(liLinPro, 19) + xlHoja1.Cells(liLinPro, 23) + xlHoja1.Cells(liLinPro, 27) + xlHoja1.Cells(liLinPro, 31) + xlHoja1.Cells(liLinPro, 35) + xlHoja1.Cells(liLinPro, 39)
                                FormatoRep liLinPro, nColCom, 4
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'EJECUTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom + 1) = "OCTUBRE EJEC. CONSOL"
                                FormatoRep liLinPro, nColCom, 1
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 1) = xlHoja1.Cells(liLinPro, 4) + xlHoja1.Cells(liLinPro, 8) + xlHoja1.Cells(liLinPro, 12) + xlHoja1.Cells(liLinPro, 16) + xlHoja1.Cells(liLinPro, 20) + xlHoja1.Cells(liLinPro, 24) + xlHoja1.Cells(liLinPro, 28) + xlHoja1.Cells(liLinPro, 32) + xlHoja1.Cells(liLinPro, 36) + xlHoja1.Cells(liLinPro, 40)
                                FormatoRep liLinPro, nColCom, 5
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'DIFERENCIA
                                xlHoja1.Cells(liLinPro - 1, nColCom + 2) = "DIFERENCIA"
                                FormatoRep liLinPro, nColCom, 2
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 2) = xlHoja1.Cells(liLinPro, 5) + xlHoja1.Cells(liLinPro, 9) + xlHoja1.Cells(liLinPro, 13) + xlHoja1.Cells(liLinPro, 17) + xlHoja1.Cells(liLinPro, 21) + xlHoja1.Cells(liLinPro, 25) + xlHoja1.Cells(liLinPro, 29) + xlHoja1.Cells(liLinPro, 33) + xlHoja1.Cells(liLinPro, 37) + xlHoja1.Cells(liLinPro, 41)
                                FormatoRep liLinPro, nColCom, 6
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'PORCENTAJE
                                xlHoja1.Cells(liLinPro - 1, nColCom + 3) = "PORCENTAJE"
                                FormatoRep liLinPro, nColCom, 3
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 3) = (xlHoja1.Cells(liLinPro, 6) + xlHoja1.Cells(liLinPro, 10) + xlHoja1.Cells(liLinPro, 14) + xlHoja1.Cells(liLinPro, 18) + xlHoja1.Cells(liLinPro, 22) + xlHoja1.Cells(liLinPro, 26) + xlHoja1.Cells(liLinPro, 30) + xlHoja1.Cells(liLinPro, 34) + xlHoja1.Cells(liLinPro, 38) + xlHoja1.Cells(liLinPro, 42)) / 10
                                FormatoRep liLinPro, nColCom, 7
                                liLinPro = liLinPro + 1
                            Next
                        Case 11: nColCom = 47 'NOVIEMBRE
                            liLinPro = 5 'PROYECTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom) = "NOVIEMBRE  PROY. CONSOL"
                                FormatoRep liLinPro, nColCom, 0
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom) = xlHoja1.Cells(liLinPro, 3) + xlHoja1.Cells(liLinPro, 7) + xlHoja1.Cells(liLinPro, 11) + xlHoja1.Cells(liLinPro, 15) + xlHoja1.Cells(liLinPro, 19) + xlHoja1.Cells(liLinPro, 23) + xlHoja1.Cells(liLinPro, 27) + xlHoja1.Cells(liLinPro, 31) + xlHoja1.Cells(liLinPro, 35) + xlHoja1.Cells(liLinPro, 39) + xlHoja1.Cells(liLinPro, 43)
                                FormatoRep liLinPro, nColCom, 4
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'EJECUTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom + 1) = "NOVIEMBRE EJEC. CONSOL"
                                FormatoRep liLinPro, nColCom, 1
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 1) = xlHoja1.Cells(liLinPro, 4) + xlHoja1.Cells(liLinPro, 8) + xlHoja1.Cells(liLinPro, 12) + xlHoja1.Cells(liLinPro, 16) + xlHoja1.Cells(liLinPro, 20) + xlHoja1.Cells(liLinPro, 24) + xlHoja1.Cells(liLinPro, 28) + xlHoja1.Cells(liLinPro, 32) + xlHoja1.Cells(liLinPro, 36) + xlHoja1.Cells(liLinPro, 40) + xlHoja1.Cells(liLinPro, 44)
                                FormatoRep liLinPro, nColCom, 5
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'DIFERENCIA
                                xlHoja1.Cells(liLinPro - 1, nColCom + 2) = "DIFERENCIA"
                                FormatoRep liLinPro, nColCom, 2
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 2) = xlHoja1.Cells(liLinPro, 5) + xlHoja1.Cells(liLinPro, 9) + xlHoja1.Cells(liLinPro, 13) + xlHoja1.Cells(liLinPro, 17) + xlHoja1.Cells(liLinPro, 21) + xlHoja1.Cells(liLinPro, 25) + xlHoja1.Cells(liLinPro, 29) + xlHoja1.Cells(liLinPro, 33) + xlHoja1.Cells(liLinPro, 37) + xlHoja1.Cells(liLinPro, 41) + xlHoja1.Cells(liLinPro, 45)
                                FormatoRep liLinPro, nColCom, 6
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'PORCENTAJE
                                xlHoja1.Cells(liLinPro - 1, nColCom + 3) = "PORCENTAJE"
                                FormatoRep liLinPro, nColCom, 3
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 3) = (xlHoja1.Cells(liLinPro, 6) + xlHoja1.Cells(liLinPro, 10) + xlHoja1.Cells(liLinPro, 14) + xlHoja1.Cells(liLinPro, 18) + xlHoja1.Cells(liLinPro, 22) + xlHoja1.Cells(liLinPro, 26) + xlHoja1.Cells(liLinPro, 30) + xlHoja1.Cells(liLinPro, 34) + xlHoja1.Cells(liLinPro, 38) + xlHoja1.Cells(liLinPro, 42) + xlHoja1.Cells(liLinPro, 46)) / 11
                                FormatoRep liLinPro, nColCom, 7
                                liLinPro = liLinPro + 1
                            Next
                        Case 12: nColCom = 51 'DICIEMBRE
                            liLinPro = 5 'PROYECTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom) = "DICIEMBRE  PROY. CONSOL"
                                FormatoRep liLinPro, nColCom, 0
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom) = xlHoja1.Cells(liLinPro, 3) + xlHoja1.Cells(liLinPro, 7) + xlHoja1.Cells(liLinPro, 11) + xlHoja1.Cells(liLinPro, 15) + xlHoja1.Cells(liLinPro, 19) + xlHoja1.Cells(liLinPro, 23) + xlHoja1.Cells(liLinPro, 27) + xlHoja1.Cells(liLinPro, 31) + xlHoja1.Cells(liLinPro, 35) + xlHoja1.Cells(liLinPro, 39) + xlHoja1.Cells(liLinPro, 43) + xlHoja1.Cells(liLinPro, 47)
                                FormatoRep liLinPro, nColCom, 4
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'EJECUTADO CONSOLIDADO
                                xlHoja1.Cells(liLinPro - 1, nColCom + 1) = "DICIEMBRE EJEC. CONSOL"
                                FormatoRep liLinPro, nColCom, 1
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 1) = xlHoja1.Cells(liLinPro, 4) + xlHoja1.Cells(liLinPro, 8) + xlHoja1.Cells(liLinPro, 12) + xlHoja1.Cells(liLinPro, 16) + xlHoja1.Cells(liLinPro, 20) + xlHoja1.Cells(liLinPro, 24) + xlHoja1.Cells(liLinPro, 28) + xlHoja1.Cells(liLinPro, 32) + xlHoja1.Cells(liLinPro, 36) + xlHoja1.Cells(liLinPro, 40) + xlHoja1.Cells(liLinPro, 44) + xlHoja1.Cells(liLinPro, 48)
                                FormatoRep liLinPro, nColCom, 5
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'DIFERENCIA
                                xlHoja1.Cells(liLinPro - 1, nColCom + 2) = "DIFERENCIA"
                                FormatoRep liLinPro, nColCom, 2
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 2) = xlHoja1.Cells(liLinPro, 5) + xlHoja1.Cells(liLinPro, 9) + xlHoja1.Cells(liLinPro, 13) + xlHoja1.Cells(liLinPro, 17) + xlHoja1.Cells(liLinPro, 21) + xlHoja1.Cells(liLinPro, 25) + xlHoja1.Cells(liLinPro, 29) + xlHoja1.Cells(liLinPro, 33) + xlHoja1.Cells(liLinPro, 37) + xlHoja1.Cells(liLinPro, 41) + xlHoja1.Cells(liLinPro, 45) + xlHoja1.Cells(liLinPro, 49)
                                FormatoRep liLinPro, nColCom, 6
                                liLinPro = liLinPro + 1
                            Next
                            liLinPro = 5 'PORCENTAJE
                                xlHoja1.Cells(liLinPro - 1, nColCom + 3) = "PORCENTAJE"
                                FormatoRep liLinPro, nColCom, 3
                            For n = 1 To nCantFila
                                xlHoja1.Cells(liLinPro, nColCom + 3) = (xlHoja1.Cells(liLinPro, 6) + xlHoja1.Cells(liLinPro, 10) + xlHoja1.Cells(liLinPro, 14) + xlHoja1.Cells(liLinPro, 18) + xlHoja1.Cells(liLinPro, 22) + xlHoja1.Cells(liLinPro, 26) + xlHoja1.Cells(liLinPro, 30) + xlHoja1.Cells(liLinPro, 34) + xlHoja1.Cells(liLinPro, 38) + xlHoja1.Cells(liLinPro, 42) + xlHoja1.Cells(liLinPro, 46) + xlHoja1.Cells(liLinPro, 50)) / 12
                                FormatoRep liLinPro, nColCom, 7
                                liLinPro = liLinPro + 1
                            Next
                End Select
                
                GeneraComentario nMES, nColCom + 4, "COMENTARIO"
            Set rs = Nothing
    PB1.value = 6
        Set RSTEMP = Nothing
    If nTipoRep = 2 Then
        Select Case nMES
            Case 1: xlHoja1.Columns("C:F").Delete Shift:=xlToLeft
            Case 2: xlHoja1.Columns("C:J").Delete Shift:=xlToLeft
            Case 3: xlHoja1.Columns("C:N").Delete Shift:=xlToLeft
            Case 4: xlHoja1.Columns("C:R").Delete Shift:=xlToLeft
            Case 5: xlHoja1.Columns("C:V").Delete Shift:=xlToLeft
            Case 6: xlHoja1.Columns("C:Z").Delete Shift:=xlToLeft
            Case 7: xlHoja1.Columns("C:AD").Delete Shift:=xlToLeft
            Case 8: xlHoja1.Columns("C:AH").Delete Shift:=xlToLeft
            Case 9: xlHoja1.Columns("C:AL").Delete Shift:=xlToLeft
            Case 10: xlHoja1.Columns("C:AP").Delete Shift:=xlToLeft
            Case 11: xlHoja1.Columns("C:AT").Delete Shift:=xlToLeft
            Case 12: xlHoja1.Columns("C:AX").Delete Shift:=xlToLeft
        End Select
        xlHoja1.Cells(1, 5) = "FECHA REPORTE"
        xlHoja1.Cells(2, 5) = gdFecSis
        xlHoja1.Cells(1, 5).Font.Bold = True
        xlHoja1.Cells(2, 5).Font.Bold = True
    End If
    PB1.value = 7
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsarchivo
        ExcelEnd App.path & "\Spooler\" & glsarchivo, xlAplicacion, xlLibro, xlHoja1
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsarchivo
        Call CargaArchivo(glsarchivo, App.path & "\SPOOLER\")
    PB1.Visible = False
    Me.lblProc.Visible = False
    Exit Sub
GeneraReporteMesErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
End Sub

Private Sub CargaDatos(ByVal nMesBal As Integer)
    Dim oRepCtaColumna As DRepCtaColumna
    Set oRepCtaColumna = New DRepCtaColumna
    Dim R As ADODB.Recordset
    Set R = New ADODB.Recordset
    Dim oRep As New DRepFormula
    Dim nReg As Integer
    nContBal = 0
    ReDim EstBal(0)
        Set R = oRepCtaColumna.GetRepProyectados(nMesBal, txtAnio.Valor)
          
          Do While Not R.EOF
              nContBal = nContBal + 1
              ReDim Preserve EstBal(nContBal)
              EstBal(nContBal - 1).cCodCta = Trim(R!nNivel)
              EstBal(nContBal - 1).cDescrip = Trim(R!cConcepto)
              EstBal(nContBal - 1).cFormula = DepuraEquivalentes(Trim(R!cFormula))
              R.MoveNext
          Loop
        RSClose R
        Set oRep = Nothing
    Set R = Nothing
End Sub

Private Sub GeneraReporteEjec(ByVal nMesBal As Integer, ByVal nCol As Integer, ByVal cCabecera As String)
Dim i As Integer
Dim k As Integer
Dim j As Integer
Dim liLinRep As Integer
Dim CTemp As String
Dim R As New ADODB.Recordset
Dim nMontoMes As Double
Dim CadSql As String
Dim CadFormula1 As String
Dim L As ListItem
Dim nFormula As New NInterpreteFormula
Dim nImporte As Currency
Dim nTipC As Currency
Dim oCtaCont As DbalanceCont
Set oCtaCont = New DbalanceCont
   
   Set oNBal = New NBalanceCont
   DoEvents
   liLinRep = 5
   xlHoja1.Cells(4, nCol) = cCabecera
   xlHoja1.Cells(4, nCol).Borders.LineStyle = 1
   xlHoja1.Cells(4, nCol).ColumnWidth = 15
   xlHoja1.Cells(4, nCol).Font.Size = 8
   
   xlHoja1.Cells(4, nCol + 1) = "DIF."
   xlHoja1.Cells(4, nCol + 1).Borders.LineStyle = 1
   xlHoja1.Cells(4, nCol + 1).ColumnWidth = 15
   xlHoja1.Cells(4, nCol + 1).Font.Size = 8
   
   xlHoja1.Cells(4, nCol + 2) = "% Ejec."
   xlHoja1.Cells(4, nCol + 2).Borders.LineStyle = 1
   xlHoja1.Cells(4, nCol + 2).ColumnWidth = 15
   xlHoja1.Cells(4, nCol + 2).Font.Size = 8
   
   dFecha = DateAdd("m", 1, CDate("01/" & Format(nMesBal, "00") & "/" & Format(txtAnio.Valor, "0000"))) - 1
   nTipC = oNBal.GetTipCambioBalance(Format(dFecha, gsFormatoMovFecha))
   
   For i = 0 To nContBal - 1
      CTemp = ""
      nCuentas = 0
      EstBal(i).cFormula = DepuraFormula(EstBal(i).cFormula)
      
      ReDim Cuentas(0)
      For k = 1 To Len(EstBal(i).cFormula)
          If UCase(Mid(EstBal(i).cFormula, k, 3)) = "DBO" Then
               CTemp = CTemp + Left(Mid(EstBal(i).cFormula, k, InStr(1, Mid(EstBal(i).cFormula, k), ")")), 3) & "." & Mid(Mid(EstBal(i).cFormula, k, InStr(1, Mid(EstBal(i).cFormula, k), ")")), 4)
               k = k + InStr(1, Mid(EstBal(i).cFormula, k), ")") - 1
          ElseIf Mid(EstBal(i).cFormula, k, 1) >= "0" And Mid(EstBal(i).cFormula, k, 1) <= "9" Then
              CTemp = CTemp + Mid(EstBal(i).cFormula, k, 1)
          Else
              If Len(CTemp) > 0 Then
                  nCuentas = nCuentas + 1
                  ReDim Preserve Cuentas(nCuentas)
                  Cuentas(nCuentas - 1).cCta = CTemp
              End If
              CTemp = ""
          End If
      Next k
      If Len(CTemp) > 0 Then
          nCuentas = nCuentas + 1
          ReDim Preserve Cuentas(nCuentas)
          Cuentas(nCuentas - 1).cCta = CTemp
      End If
      'Carga Valores de las Cuentas
      For k = 0 To nCuentas - 1
        If UCase(Left(Cuentas(k).cCta, 4)) = "DBO." Then
           'Saldo por cuenta
           nMontoMes = nFormula.EjecutaFuncion(Cuentas(k).cCta)
        Else
           Dim psMoneda As String
           If Len(Cuentas(k).cCta) > 3 Then
              psMoneda = Mid(Cuentas(k).cCta, 3, 1)
           Else
              psMoneda = 0
           End If
           nMontoMes = oCtaCont.ObtenerCtaContBalanceMensual(Cuentas(k).cCta, CDate(dFecha), psMoneda, nTipC)
        End If
        'Actualiza Montos
        Cuentas(k).nMES = nMontoMes
      Next k
      'Genero las 3 formulas para las 3 monedas
      CTemp = ""
      CadFormula1 = ""
                  
        For k = 1 To Len(EstBal(i).cFormula)
            If UCase(Mid(EstBal(i).cFormula, k, 3)) = "DBO" Then
                 CTemp = CTemp + Left(Mid(EstBal(i).cFormula, k, InStr(1, Mid(EstBal(i).cFormula, k), ")")), 3) & "." & Mid(Mid(EstBal(i).cFormula, k, InStr(1, Mid(EstBal(i).cFormula, k), ")")), 4)
                 k = k + InStr(1, Mid(EstBal(i).cFormula, k), ")") - 1
            ElseIf (Mid(EstBal(i).cFormula, k, 1) >= "0" And Mid(EstBal(i).cFormula, k, 1) <= "9") Or (Mid(EstBal(i).cFormula, k, 1) = ".") Then
                CTemp = CTemp + Mid(EstBal(i).cFormula, k, 1)
            ElseIf Mid(EstBal(i).cFormula, k, 1) = "$" Then
                 j = InStr(k + 1, EstBal(i).cFormula, "$") - 1
                 nImporte = Mid(EstBal(i).cFormula, k + 1, j - k)
                 CadFormula1 = CadFormula1 + Format(nImporte, gsFormatoNumeroDato)
                 k = j + 1
            Else
                 If Len(CTemp) > 0 Then
                     'busca su equivalente en monto
                     For j = 0 To nCuentas
                         If Cuentas(j).cCta = CTemp Then
                             CadFormula1 = CadFormula1 + Format(Cuentas(j).nMES, gsFormatoNumeroDato)
                             Exit For
                         End If
                     Next j
                 End If
                 CTemp = ""
                 CadFormula1 = CadFormula1 + Mid(EstBal(i).cFormula, k, 1)
            End If
        Next k
         
         If Len(CTemp) > 0 Then
             'busca su equivalente en monto
             For j = 0 To nCuentas
                 If Cuentas(j).cCta = CTemp Then
                     CadFormula1 = CadFormula1 + Format(Cuentas(j).nMES, gsFormatoNumeroDato)
                     Exit For
                 End If
             Next j
         End If
         nMontoMes = 0
         
         nMontoMes = nFormula.ExprANum(CadFormula1, EstBal(i).cCodCta)
         
        'Asignacion de Valores
        xlHoja1.Cells(liLinRep, nCol) = nMontoMes
        xlHoja1.Cells(liLinRep, nCol).NumberFormat = "#,###0.00"
        xlHoja1.Cells(liLinRep, nCol).Borders.LineStyle = 1
        xlHoja1.Cells(liLinRep, nCol).ColumnWidth = 15
        xlHoja1.Cells(liLinRep, nCol).Font.Size = 8
        
        xlHoja1.Cells(liLinRep, nCol + 1) = xlHoja1.Cells(liLinRep, nCol) - xlHoja1.Cells(liLinRep, nCol - 1)
        xlHoja1.Cells(liLinRep, nCol + 1).NumberFormat = "#,###0.00"
        xlHoja1.Cells(liLinRep, nCol + 1).Borders.LineStyle = 1
        xlHoja1.Cells(liLinRep, nCol + 1).ColumnWidth = 15
        xlHoja1.Cells(liLinRep, nCol + 1).Font.Size = 8
        
        Dim nDiv As Currency
        If xlHoja1.Cells(liLinRep, nCol - 1) = 0 Then
            nDiv = 0
        Else
            nDiv = (xlHoja1.Cells(liLinRep, nCol) / xlHoja1.Cells(liLinRep, nCol - 1))
        End If
        xlHoja1.Cells(liLinRep, nCol + 2) = nDiv
        xlHoja1.Cells(liLinRep, nCol + 2).NumberFormat = "#,###0.00"
        xlHoja1.Cells(liLinRep, nCol + 2).Borders.LineStyle = 1
        xlHoja1.Cells(liLinRep, nCol + 2).ColumnWidth = 15
        xlHoja1.Cells(liLinRep, nCol + 2).Font.Size = 8
        liLinRep = liLinRep + 1
   Next i
Set nFormula = Nothing
End Sub

Private Sub GeneraComentario(ByVal nMesBal As Integer, ByVal nCol As Integer, ByVal cCabecera As String)
Dim n As Integer
Dim i As Integer
Dim k As Integer
Dim j As Integer
Dim liLinRep As Integer
Dim CTemp As String
Dim sFechaRep As String
Dim CadSql As String
Dim CadFormula1 As String
Dim sGlosa1 As String
Dim sGlosa2 As String
Dim sGlosaA As String
Dim sGlosaB As String
Dim sGlosaPivotA As String
Dim sGlosaPivotB As String
Dim sGlosaFinal1 As String
Dim sGlosaFinal2 As String
Dim sGlosaFinal As String
Dim nImporte As Currency
Dim nTipC As Currency
Dim nMontoGlosa1 As Currency
Dim nMontoGlosa2 As Currency
Dim nMontoGlosaA As Currency
Dim nMontoGlosaB As Currency
Dim nMontoA As Currency
Dim nMontoB As Currency
Dim nMontoMes As Double
Dim L As ListItem
Dim R As New ADODB.Recordset
Dim nFormula As New NInterpreteFormula
Dim oRepCtaColumna As DRepCtaColumna
Set oRepCtaColumna = New DRepCtaColumna
Dim oCtaCont As DbalanceCont
Set oCtaCont = New DbalanceCont
   
   Set oNBal = New NBalanceCont
   DoEvents
   liLinRep = 5
   xlHoja1.Cells(4, nCol) = cCabecera
   xlHoja1.Cells(4, nCol).Borders.LineStyle = 1
   xlHoja1.Cells(4, nCol).ColumnWidth = 15
   xlHoja1.Cells(4, nCol).Font.Size = 8
   
   dFecha = DateAdd("m", 1, CDate("01/" & Format(nMesBal, "00") & "/" & Format(txtAnio.Valor, "0000"))) - 1
   sFechaRep = Format(dFecha, "yyyyMM")
   
   For i = 0 To nContBal - 1
      CTemp = ""
      nCuentas = 0
      EstBal(i).cFormula = DepuraFormula(EstBal(i).cFormula)
      
      ReDim Cuentas(0)
      For k = 1 To Len(EstBal(i).cFormula)
          If UCase(Mid(EstBal(i).cFormula, k, 3)) = "DBO" Then
               CTemp = CTemp + Left(Mid(EstBal(i).cFormula, k, InStr(1, Mid(EstBal(i).cFormula, k), ")")), 3) & "." & Mid(Mid(EstBal(i).cFormula, k, InStr(1, Mid(EstBal(i).cFormula, k), ")")), 4)
               k = k + InStr(1, Mid(EstBal(i).cFormula, k), ")") - 1
          ElseIf Mid(EstBal(i).cFormula, k, 1) >= "0" And Mid(EstBal(i).cFormula, k, 1) <= "9" Then
              CTemp = CTemp + Mid(EstBal(i).cFormula, k, 1)
          Else
              If Len(CTemp) > 0 Then
                  nCuentas = nCuentas + 1
                  ReDim Preserve Cuentas(nCuentas)
                  Cuentas(nCuentas - 1).cCta = CTemp
              End If
              CTemp = ""
          End If
      Next k
      If Len(CTemp) > 0 Then
          nCuentas = nCuentas + 1
          ReDim Preserve Cuentas(nCuentas)
          Cuentas(nCuentas - 1).cCta = CTemp
      End If
      'Carga Valores de las Cuentas
      For k = 0 To nCuentas - 1
        If UCase(Left(Cuentas(k).cCta, 4)) = "DBO." Then
           'Saldo por cuenta
           nMontoMes = nFormula.EjecutaFuncion(Cuentas(k).cCta)
        Else
           'codigo para la glosa
           Dim psMoneda As String
           If Len(Cuentas(k).cCta) > 3 Then
              If Mid(Cuentas(k).cCta, 3, 1) = 0 Then
                psMoneda = "[12]"
              Else
                psMoneda = Mid(Cuentas(k).cCta, 3, 1)
              End If
              Set R = oRepCtaColumna.GetMovGlosa(sFechaRep, Left(Cuentas(k).cCta, 2) + psMoneda + Mid(Cuentas(k).cCta, 4, Len(Cuentas(k).cCta) - 3))
           Else
              psMoneda = ""
              Set R = oRepCtaColumna.GetMovGlosa(sFechaRep, Left(Cuentas(k).cCta, 2) + psMoneda)
           End If
           Do Until R.EOF
               If n = 0 Then
                   sGlosa1 = R!cMovDesc
                   nMontoGlosa1 = R!nMovImporte
               Else
                   sGlosa2 = R!cMovDesc
                   nMontoGlosa2 = R!nMovImporte
               End If
               n = n + 1
               R.MoveNext
           Loop
           '********************
        End If
        n = 0
        'Actualiza Comentario
            If nCuentas = 1 Then
                If R.RecordCount > 0 Then
                    Cuentas(k).cDescrip = sGlosa1 & "; " & sGlosa2
                Else
                    Cuentas(k).cDescrip = ""
                End If
            Else
                If k = 0 Then
                    If nMontoGlosa1 > nMontoGlosa2 Then
                        sGlosaA = sGlosa1
                        sGlosaB = sGlosa2
                        nMontoA = nMontoGlosa1
                        nMontoB = nMontoGlosa2
                    Else
                        sGlosaA = sGlosa2
                        sGlosaB = sGlosa1
                        nMontoA = nMontoGlosa2
                        nMontoB = nMontoGlosa1
                    End If
                Else
                    If nMontoGlosa1 > nMontoGlosa2 Then
                        sGlosaPivotA = sGlosa1
                        sGlosaPivotB = sGlosa2
                        nMontoGlosaA = nMontoGlosa1
                        nMontoGlosaB = nMontoGlosa2
                    Else
                        sGlosaPivotA = sGlosa2
                        sGlosaPivotB = sGlosa1
                        nMontoGlosaA = nMontoGlosa2
                        nMontoGlosaB = nMontoGlosa1
                    End If
                    If nMontoA > nMontoGlosaA And nMontoB > nMontoGlosaA Then
                        sGlosaA = sGlosaA
                        sGlosaB = sGlosaB
                        nMontoA = nMontoA
                        nMontoB = nMontoB
                    ElseIf nMontoA > nMontoGlosaA And nMontoB < nMontoGlosaA Then
                        sGlosaA = sGlosaA
                        sGlosaB = nMontoGlosaA
                        nMontoA = nMontoA
                        nMontoB = nMontoGlosaA
                    ElseIf nMontoGlosaA > nMontoA And nMontoGlosaB > nMontoA Then
                        sGlosaA = sGlosaPivotA
                        sGlosaB = sGlosaPivotB
                        nMontoA = nMontoGlosaA
                        nMontoB = nMontoGlosaB
                    ElseIf nMontoGlosaA > nMontoA And nMontoGlosaB < nMontoA Then
                        sGlosaB = sGlosaA
                        sGlosaA = sGlosaPivotA
                        nMontoB = nMontoA
                        nMontoA = nMontoGlosaA
                    End If
                End If
                Cuentas(k).cDescrip = sGlosaA + "; " + sGlosaB
                sGlosa1 = ""
                sGlosa2 = ""
                nMontoGlosa1 = 0
                nMontoGlosa2 = 0
            End If
      Next k
      'Genero las 3 formulas para las 3 monedas
      CTemp = ""
        For k = 1 To Len(EstBal(i).cFormula)
            If UCase(Mid(EstBal(i).cFormula, k, 3)) = "DBO" Then
                 CTemp = CTemp + Left(Mid(EstBal(i).cFormula, k, InStr(1, Mid(EstBal(i).cFormula, k), ")")), 3) & "." & Mid(Mid(EstBal(i).cFormula, k, InStr(1, Mid(EstBal(i).cFormula, k), ")")), 4)
                 k = k + InStr(1, Mid(EstBal(i).cFormula, k), ")") - 1
            ElseIf (Mid(EstBal(i).cFormula, k, 1) >= "0" And Mid(EstBal(i).cFormula, k, 1) <= "9") Or (Mid(EstBal(i).cFormula, k, 1) = ".") Then
                CTemp = CTemp + Mid(EstBal(i).cFormula, k, 1)
            ElseIf Mid(EstBal(i).cFormula, k, 1) = "$" Then
                 j = InStr(k + 1, EstBal(i).cFormula, "$") - 1
                 nImporte = Mid(EstBal(i).cFormula, k + 1, j - k)
                 CadFormula1 = CadFormula1 + Format(nImporte, gsFormatoNumeroDato)
                 k = j + 1
            Else
                 If Len(CTemp) > 0 Then
                     'busca su equivalente en monto
                     For j = 0 To nCuentas
                         If Cuentas(j).cCta = CTemp Then
                             CadFormula1 = Cuentas(j).cDescrip
                             Exit For
                         End If
                     Next j
                 End If
                 CTemp = ""
                 CadFormula1 = CadFormula1 + Mid(EstBal(i).cFormula, k, 1)
            End If
        Next k
         
         If Len(CTemp) > 0 Then
             For j = 0 To nCuentas
                 If Cuentas(j).cCta = CTemp Then
                     CadFormula1 = Cuentas(j).cDescrip
                     Exit For
                 End If
             Next j
         End If
         
         Set R = oRepCtaColumna.GetDatosProyEjec(nMesBal, txtAnio.Valor)
         Do Until R.EOF
            If R!nGlosa = 1 And R!cFormula = EstBal(i).cFormula Then
                sGlosaFinal = CadFormula1
            End If
            R.MoveNext
         Loop
         sGlosa1 = ""
         nMontoGlosa1 = 0
         sGlosa2 = ""
         nMontoGlosa2 = 0
         CadFormula1 = ""
         Set R = Nothing
        'Asignacion de Valores
        xlHoja1.Cells(liLinRep, nCol) = sGlosaFinal
        xlHoja1.Cells(liLinRep, nCol).Borders.LineStyle = 1
        xlHoja1.Cells(liLinRep, nCol).ColumnWidth = 50
        xlHoja1.Cells(liLinRep, nCol).Font.Size = 8
        liLinRep = liLinRep + 1
        sGlosaFinal = ""
   Next i
Set nFormula = Nothing
End Sub

Private Sub FormatoRep(ByVal nLin As Integer, ByVal nCol As Integer, ByVal nTpoRep As Integer)
If nTpoRep = 0 Then
    xlHoja1.Cells(nLin - 1, nCol).NumberFormat = "#,###0.00"
    xlHoja1.Cells(nLin - 1, nCol).Borders.LineStyle = 1
    xlHoja1.Cells(nLin - 1, nCol).ColumnWidth = 15
    xlHoja1.Cells(nLin - 1, nCol).Font.Size = 8
ElseIf nTpoRep = 1 Then
    xlHoja1.Cells(nLin - 1, nCol + 1).NumberFormat = "#,###0.00"
    xlHoja1.Cells(nLin - 1, nCol + 1).Borders.LineStyle = 1
    xlHoja1.Cells(nLin - 1, nCol + 1).ColumnWidth = 15
    xlHoja1.Cells(nLin - 1, nCol + 1).Font.Size = 8
ElseIf nTpoRep = 2 Then
    xlHoja1.Cells(nLin - 1, nCol + 2).NumberFormat = "#,###0.00"
    xlHoja1.Cells(nLin - 1, nCol + 2).Borders.LineStyle = 1
    xlHoja1.Cells(nLin - 1, nCol + 2).ColumnWidth = 15
    xlHoja1.Cells(nLin - 1, nCol + 2).Font.Size = 8
ElseIf nTpoRep = 3 Then
    xlHoja1.Cells(nLin - 1, nCol + 3).NumberFormat = "#,###0.00"
    xlHoja1.Cells(nLin - 1, nCol + 3).Borders.LineStyle = 1
    xlHoja1.Cells(nLin - 1, nCol + 3).ColumnWidth = 15
    xlHoja1.Cells(nLin - 1, nCol + 3).Font.Size = 8
ElseIf nTpoRep = 4 Then
    xlHoja1.Cells(nLin, nCol).NumberFormat = "#,###0.00"
    xlHoja1.Cells(nLin, nCol).Borders.LineStyle = 1
    xlHoja1.Cells(nLin, nCol).ColumnWidth = 15
    xlHoja1.Cells(nLin, nCol).Font.Size = 8
ElseIf nTpoRep = 5 Then
    xlHoja1.Cells(nLin, nCol + 1).NumberFormat = "#,###0.00"
    xlHoja1.Cells(nLin, nCol + 1).Borders.LineStyle = 1
    xlHoja1.Cells(nLin, nCol + 1).ColumnWidth = 15
    xlHoja1.Cells(nLin, nCol + 1).Font.Size = 8
ElseIf nTpoRep = 6 Then
    xlHoja1.Cells(nLin, nCol + 2).NumberFormat = "#,###0.00"
    xlHoja1.Cells(nLin, nCol + 2).Borders.LineStyle = 1
    xlHoja1.Cells(nLin, nCol + 2).ColumnWidth = 15
    xlHoja1.Cells(nLin, nCol + 2).Font.Size = 8
ElseIf nTpoRep = 7 Then
    xlHoja1.Cells(nLin, nCol + 3).NumberFormat = "#,###0.00"
    xlHoja1.Cells(nLin, nCol + 3).Borders.LineStyle = 1
    xlHoja1.Cells(nLin, nCol + 3).ColumnWidth = 15
    xlHoja1.Cells(nLin, nCol + 3).Font.Size = 8
End If
End Sub

Private Function DepuraEquivalentes(psEquival As String) As String
Dim j As Integer
Dim CadTemp As String
   CadTemp = ""
   For j = 1 To Len(psEquival)
       If Mid(psEquival, j, 1) <> "." Then
           CadTemp = CadTemp + Mid(psEquival, j, 1)
       End If
   Next j
   DepuraEquivalentes = CadTemp
End Function

Private Function DepuraFormula(sFormula As String) As String
Dim sCad As String
Dim R As New ADODB.Recordset
Dim sSql As String
Dim i As Integer
Dim sCadRes As String
Dim bFinal As Boolean
Dim sCod As String
    sCad = sFormula
    i = 1
    sCadRes = ""
    Do While i <= Len(sCad)
        If Mid(sCad, i, 1) <> "#" Then
            sCadRes = sCadRes + Mid(sCad, i, 1)
        Else
            i = i + 2
            bFinal = False
            sCod = ""
            Do While Not bFinal
                If Mid(sCad, i, 1) <> "]" Then
                    sCod = sCod + Mid(sCad, i, 1)
                Else
                    bFinal = True
                End If
                i = i + 1
            Loop
            sCadRes = sCadRes + DepuraMichi(sCod)
            i = i - 1
        End If
        i = i + 1
    Loop
    DepuraFormula = sCadRes
End Function

Private Function DepuraMichi(sCodigo As String) As String
Dim R As New ADODB.Recordset
Dim sSql As String
Dim sCadFor As String
Dim sCadRes As String
Dim i As Integer
Dim bFinal As Boolean
Dim sCod As String
Dim TodoASoles As Boolean
Dim Aspersand As Boolean
Dim oRepFormula As New DRepFormula
    TodoASoles = False
    Aspersand = False
    If Mid(sCodigo, 1, 1) = "&" Then
        sCodigo = Mid(sCodigo, 2, Len(sCodigo) - 1)
        TodoASoles = True
        Aspersand = True
    End If
    Set R = oRepFormula.CargaRepFormula(sCodigo, gsOpeCod)
        sCadFor = Trim(R!cFormula)
    R.Close
    sCadRes = ""
    i = 1
    Do While i <= Len(sCadFor)
        If Mid(sCadFor, i, 1) <> "#" Then
            If TodoASoles And Mid(sCadFor, i, 1) >= "0" And Mid(sCadFor, i, 1) <= "9" And Aspersand Then
                sCadRes = sCadRes + "&"
                Aspersand = False
            Else
                If Not (Mid(sCadFor, i, 1) >= "0" And Mid(sCadFor, i, 1) <= "9") Then
                    Aspersand = True
                End If
            End If
            sCadRes = sCadRes + Mid(sCadFor, i, 1)
        Else
            i = i + 2
            bFinal = False
            sCod = ""
            Do While Not bFinal
                If Mid(sCadFor, i, 1) <> "]" Then
                    sCod = sCod + Mid(sCadFor, i, 1)
                Else
                    bFinal = True
                End If
                i = i + 1
            Loop
            sCadRes = sCadRes + DepuraMichi(sCod)
            i = i - 1
        End If
        i = i + 1
    Loop
    DepuraMichi = sCadRes
End Function

