VERSION 5.00
Begin VB.Form frmMntLineasCredito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Lineas de Credito "
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12750
   Icon            =   "frmMntLineasCredito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReporteIF 
      Caption         =   "Reporte IF"
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboLinea 
      Height          =   315
      Left            =   8400
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   8760
      TabIndex        =   6
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   10080
      TabIndex        =   5
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   11400
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton CmdMostrar 
      Caption         =   "Mostrar"
      Height          =   375
      Left            =   11280
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin Sicmact.FlexEdit FECofide 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   9551
      Cols0           =   11
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Item-Persona-Crédito-Moneda-Plazo-CodProducto-Producto-Tasa-Sald.Cap-LineaCredito-LineaNueva"
      EncabezadosAnchos=   "500-3000-2000-800-700-0-1200-1200-1200-0-1500"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-10"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-1"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C-C-C-C-L-R-R-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-4-4-0-0"
      TextArray0      =   "Item"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label lblLinea 
      Caption         =   "Tipo Linea:"
      Height          =   255
      Left            =   7320
      TabIndex        =   11
      Top             =   165
      Width           =   975
   End
   Begin VB.Label lblTipoCambio 
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de cambio:"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label lblFechaReporte 
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Información del Reporte al:"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   6240
      Width           =   2055
   End
End
Attribute VB_Name = "frmMntLineasCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
'*****************************ALPA 20110808****************************************
'**********************************************************************************
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub cmdActualizar_Click()
Dim oCredito As DCreditos
Dim i As Integer

If Trim(cboLinea.Text) = "" Then
        MsgBox "Debe seleccionar la linea de credito ", vbCritical
        Exit Sub
    End If
If (FECofide.Rows - 1) <= 0 Then
    MsgBox "Debe actualizar los datos de la lista de créditos", vbCritical
    Exit Sub
End If

Set oCredito = New DCreditos
    For i = 1 To FECofide.Rows - 1
        Call oCredito.ActualizaLineaCredito(lblFechaReporte.Caption, FECofide.TextMatrix(i, 2), FECofide.TextMatrix(i, 10), FECofide.TextMatrix(i, 9))
    Next i
    Call cmdMostrar_Click
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Actualizo la Operación "
                Set objPista = Nothing
                '****
    If MsgBox("El proceso se realizó con exito, Desea Imprimir el reporte ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Call Reporte
End Sub

Private Sub cmdMostrar_Click()
    Dim oCredito As DCreditos
    Dim oRs As ADODB.Recordset
    Dim i As Integer
    Set oCredito = New DCreditos
    Set oRs = New ADODB.Recordset
    
    'ALPA 20120601
    If Trim(cboLinea.Text) = "" Then
        MsgBox "Debe seleccionar la linea de credito ", vbCritical
        Exit Sub
    End If
    
'    right(cboLinea.Text,5)
    Set oRs = oCredito.RecuperaDatosParaMntCreditosCofide(Trim(Right(cboLinea.Text, 5)))
    
    LimpiaFlex FECofide
    i = 1
    Do While Not oRs.EOF
        FECofide.AdicionaFila
        FECofide.TextMatrix(oRs.Bookmark, 1) = oRs!cPersNombre
        FECofide.TextMatrix(oRs.Bookmark, 2) = oRs!cCtaCod
        FECofide.TextMatrix(oRs.Bookmark, 3) = oRs!cMoneda
        FECofide.TextMatrix(oRs.Bookmark, 4) = oRs!cPlazo
        FECofide.TextMatrix(oRs.Bookmark, 5) = oRs!cTpoCredCod
        FECofide.TextMatrix(oRs.Bookmark, 6) = oRs!cTpoCredDesc
        FECofide.TextMatrix(oRs.Bookmark, 7) = oRs!nTasaInteres
        FECofide.TextMatrix(oRs.Bookmark, 8) = oRs!nMontoApr
        FECofide.TextMatrix(oRs.Bookmark, 9) = oRs!cLineaCred
        FECofide.TextMatrix(oRs.Bookmark, 10) = oRs!cLineaCredNew

        oRs.MoveNext
    Loop
    
    Set oCredito = Nothing
End Sub

Private Sub cmdImprimir_Click()
    If Trim(cboLinea.Text) = "" Then
        MsgBox "Debe seleccionar la linea de credito ", vbCritical
        Exit Sub
    End If
    Call Reporte
End Sub

'Private Sub cmdReporteIF_Click()
'Dim fs As Scripting.FileSystemObject
'    Dim lbExisteHoja As Boolean
'    Dim lsArchivo1 As String
'    Dim lsNomHoja  As String
'    Dim lsNombreAgencia As String
'    Dim lsCodAgencia As String
'    Dim lsMes As String
'    Dim lnContador As Integer
'    Dim lsArchivo As String
'    Dim xlsAplicacion As Excel.Application
'    Dim xlsLibro As Excel.Workbook
'    Dim xlHoja1 As Excel.Worksheet
'
'    Dim rsCreditos As ADODB.Recordset
'
'    Dim oCreditos As New DCreditos
'
'    Dim sTexto As String
'    Dim sDocFecha As String
'    Dim nSaltoContador As Double
'    Dim sFecha As String
'    Dim sMov As String
'    Dim sDoc As String
'    Dim n As Integer
'    Dim pnLinPage As Integer
'    Dim nMES As Integer
'    Dim nSaldo12 As Currency
'    Dim nContTotal As Double
'    Dim nPase As Integer
'    Dim lnSaldoVigente, lnSaldoRefinanciado, lnSaldoVencido, lnSaldoJudicial As Currency
'    Dim lnCalGen0, lnCalGen1, lnCalGen2, lnCalGen3, lnCalGen4 As Currency
'    Dim lnSaldoAdeudados As Currency
'    Set fs = New Scripting.FileSystemObject
'    Set xlsAplicacion = New Excel.Application
'    lsArchivo = "ReporteMIF"
'    'Primera Hoja ******************************************************
'    lsNomHoja = "ReporteMIF"
'    '*******************************************************************
'    lsArchivo1 = "\spooler\reporteMIF" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Trim(Left(cboLinea.Text, 8)) & ".xls" '& Format$(Time(), "HHMMSS") & ".xls"
'    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
'        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
'    Else
'        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
'        Exit Sub
'    End If
'
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    nSaltoContador = 8
'
'    'nMES = cboMes.ListIndex + 1
'    Set rsCreditos = oCreditos.reporteMIF(lblFechaReporte.Caption, CDbl(lblTipoCambio.Caption))
'    nPase = 1
'    If (rsCreditos Is Nothing) Then
'        nPase = 0
'    End If
'    'ALPA
'
'    xlHoja1.Cells(9, 2) = Format(lblFechaReporte.Caption, "YYYY/MM/DD")
'     If nPase = 1 Then
'        Do While Not rsCreditos.EOF
'    '        DoEvents
''                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 7)).Borders.LineStyle = 1
'            lnSaldoVigente = lnSaldoVigente + rsCreditos!nSaldoVigente
'            lnSaldoRefinanciado = lnSaldoRefinanciado + rsCreditos!nSaldoRefinanciado
'            lnSaldoVencido = lnSaldoVencido + rsCreditos!nSaldoVencido
'            lnSaldoJudicial = lnSaldoJudicial + rsCreditos!nSaldoJudicial
'            If rsCreditos!cCalGen = 0 Then
'                    lnCalGen0 = rsCreditos!nSaldoVigente + rsCreditos!nSaldoRefinanciado + rsCreditos!nSaldoVencido + rsCreditos!nSaldoJudicial
'            ElseIf rsCreditos!cCalGen = 1 Then
'                    lnCalGen1 = rsCreditos!nSaldoVigente + rsCreditos!nSaldoRefinanciado + rsCreditos!nSaldoVencido + rsCreditos!nSaldoJudicial
'            ElseIf rsCreditos!cCalGen = 2 Then
'                    lnCalGen2 = rsCreditos!nSaldoVigente + rsCreditos!nSaldoRefinanciado + rsCreditos!nSaldoVencido + rsCreditos!nSaldoJudicial
'            ElseIf rsCreditos!cCalGen = 3 Then
'                    lnCalGen3 = rsCreditos!nSaldoVigente + rsCreditos!nSaldoRefinanciado + rsCreditos!nSaldoVencido + rsCreditos!nSaldoJudicial
'            ElseIf rsCreditos!cCalGen = 4 Then
'                    lnCalGen4 = rsCreditos!nSaldoVigente + rsCreditos!nSaldoRefinanciado + rsCreditos!nSaldoVencido + rsCreditos!nSaldoJudicial
'            End If
'            lnSaldoAdeudados = IIf(IsNull(rsCreditos!nSaldoVigente), 0, rsCreditos!nSaldoVigente)
'            rsCreditos.MoveNext
'            nContTotal = nContTotal + 1
'            If rsCreditos.EOF Then
'               Exit Do
'            End If
'        Loop
'        xlHoja1.Cells(14, 2) = Format(lnSaldoVigente, "#,###,###,###,##0.00")
'        xlHoja1.Cells(15, 2) = 0#
'        xlHoja1.Cells(16, 2) = Format(lnSaldoRefinanciado, "#,###,###,###,##0.00")
'        xlHoja1.Cells(17, 2) = Format(lnSaldoVencido, "#,###,###,###,##0.00")
'        xlHoja1.Cells(18, 2) = Format(lnSaldoJudicial, "#,###,###,###,##0.00")
'
'        xlHoja1.Cells(19, 2) = Format(lnSaldoAdeudados, "#,###,###,###,##0.00")
'
'        xlHoja1.Cells(22, 2) = Format(lnCalGen0, "#,###,###,###,##0.00")
'        xlHoja1.Cells(23, 2) = Format(lnCalGen1, "#,###,###,###,##0.00")
'        xlHoja1.Cells(24, 2) = Format(lnCalGen2, "#,###,###,###,##0.00")
'        xlHoja1.Cells(25, 2) = Format(lnCalGen3, "#,###,###,###,##0.00")
'        xlHoja1.Cells(26, 2) = Format(lnCalGen4, "#,###,###,###,##0.00")
'
'   End If
'    'ALPA FIN
'    Set oCreditos = Nothing
'    If nPase = 1 Then
'        rsCreditos.Close
'    End If
'    Set rsCreditos = Nothing
'
'    xlHoja1.SaveAs App.path & lsArchivo1
'    xlsAplicacion.Visible = True
'    xlsAplicacion.Windows(1).Visible = True
'    Set xlsAplicacion = Nothing
'    Set xlsLibro = Nothing
'    Set xlHoja1 = Nothing
'
'Exit Sub
'End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub FECofide_RowColChange()
If FECofide.col = 10 Then
    Dim cMoneda As String
    Dim cPlazo As String
    Dim oCredito As DCreditos
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Set oCredito = New DCreditos
    FECofide.lbEditarFlex = True
    FECofide.SetFocus
    If FECofide.TextMatrix(FECofide.row, 3) = "MN" Then
        cMoneda = "1"
    Else
        cMoneda = "2"
    End If
    
    If FECofide.TextMatrix(FECofide.row, 4) = "CP" Then
        cPlazo = "1"
    Else
        cPlazo = "2"
    End If
    Set oRs = oCredito.RecuperaDatosParaMntCreditosCofidexCredito(cMoneda, cPlazo, FECofide.TextMatrix(FECofide.row, 5), FECofide.TextMatrix(FECofide.row, 7))
    FECofide.TipoBusqueda = BuscaGrid
    FECofide.lbUltimaInstancia = True
    FECofide.AutoAdd = True
    FECofide.rsTextBuscar = oRs
End If
End Sub

Private Sub Form_Load()
Dim loTipCambio As nTipoCambio

    Set loTipCambio = New nTipoCambio
        lblTipoCambio.Caption = Format(loTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "0.###")
    Set loTipCambio = Nothing


    Dim loConstS As DConstSistemas
    CentraForm Me
    Set loConstS = New DConstSistemas
    lblFechaReporte.Caption = CDate(loConstS.LeeConstSistema(gConstSistCierreMesNegocio))
    Call InicializarLinea 'ALPA 20120601
End Sub
'ALPA 20120601
Private Sub InicializarLinea()
cboLinea.Clear
cboLinea.AddItem "COFIDE" & Space(200) & "02"
cboLinea.AddItem "FONCODES" & Space(200) & "04"
End Sub

Public Sub Reporte()
    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja  As String
    Dim lsNombreAgencia As String
    Dim lsCodAgencia As String
    Dim lsMes As String
    Dim lnContador As Integer
    Dim lsArchivo As String
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    
    Dim rsCreditos As ADODB.Recordset
    
    Dim oCreditos As New DCreditos
    
    Dim sTexto As String
    Dim sDocFecha As String
    Dim nSaltoContador As Double
    Dim sFecha As String
    Dim sMov As String
    Dim sDoc As String
    Dim n As Integer
    Dim pnLinPage As Integer
    Dim nMes As Integer
    Dim nSaldo12 As Currency
    Dim nContTotal As Double
    Dim nPase As Integer
'On Error GoTo GeneraExcelErr

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "LineaCredito"
    'Primera Hoja ******************************************************
    lsNomHoja = "LineaCredito"
    '*******************************************************************
    'MIOL 20121024, SEGUN RQ12338: SE CAMBIO "Format$(Time(), "HHMMSS")" POR "Trim(Left(cboLinea.Text, 8))" ******************
    lsArchivo1 = "\spooler\LineaCredito" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Trim(Left(cboLinea.Text, 8)) & ".xlsx"  'Format$(Time(), "HHMMSS") & ".xls"
    'END MIOL ****************************************************************************************************************
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xlsx")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    
    nSaltoContador = 8
     
    'nMES = cboMes.ListIndex + 1
    Set rsCreditos = oCreditos.ReporteLineaCredito(lblFechaReporte.Caption, Trim(Right(cboLinea.Text, 5)))
    nPase = 1
    If (rsCreditos Is Nothing) Then
        nPase = 0
    End If
     'lblTipoCambio.Caption
    xlHoja1.Cells(5, 1) = Format(lblFechaReporte.Caption, "DD") & " DE " & UCase(Format(lblFechaReporte.Caption, "MMMM")) & " DEL  " & Format(lblFechaReporte.Caption, "YYYY")
    If nPase = 1 Then
        Do While Not rsCreditos.EOF
    '        DoEvents
                xlHoja1.Range(xlHoja1.Cells(nSaltoContador, 1), xlHoja1.Cells(nSaltoContador, 7)).Borders.LineStyle = 1
                xlHoja1.Cells(nSaltoContador, 1) = rsCreditos!cCtaCod
                xlHoja1.Cells(nSaltoContador, 2) = rsCreditos!cTpoCredDesc
                xlHoja1.Cells(nSaltoContador, 3) = rsCreditos!cPersCod
                xlHoja1.Cells(nSaltoContador, 4) = rsCreditos!cPersNombre
                xlHoja1.Cells(nSaltoContador, 5) = Format(IIf(Mid(rsCreditos!cCtaCod, 9, 1) = "1", 1, Val(lblTipoCambio.Caption)) * rsCreditos!nSaldoCap, "###0.00")
                xlHoja1.Cells(nSaltoContador, 6) = rsCreditos!cLineaCred
                xlHoja1.Cells(nSaltoContador, 7) = rsCreditos!cLineaCredAnterior
                nSaltoContador = nSaltoContador + 1
            rsCreditos.MoveNext
            nContTotal = nContTotal + 1
            If rsCreditos.EOF Then
               Exit Do
            End If
        Loop
    End If
    Set oCreditos = Nothing
    If nPase = 1 Then
        rsCreditos.Close
    End If
    Set rsCreditos = Nothing
    
    xlHoja1.SaveAs App.path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", Me.Caption & " Se Genero Reporte "
                Set objPista = Nothing
                '****
Exit Sub
End Sub
