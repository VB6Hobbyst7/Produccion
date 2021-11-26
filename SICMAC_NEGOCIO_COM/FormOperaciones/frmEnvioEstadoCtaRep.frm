VERSION 5.00
Begin VB.Form frmEnvioEstadoCtaRep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generación de Envío de Estado de Cuenta Físico"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   Icon            =   "frmEnvioEstadoCtaRep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7080
      TabIndex        =   14
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CheckBox chkTodos 
      Caption         =   "Todos"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   " Filtro "
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   9495
      Begin VB.TextBox txtAnio 
         Height          =   300
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   1
         Top             =   480
         Width           =   915
      End
      Begin VB.TextBox txtAgencia 
         Height          =   320
         Left            =   5640
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   8160
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmEnvioEstadoCtaRep.frx":030A
         Left            =   360
         List            =   "frmEnvioEstadoCtaRep.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "frmEnvioEstadoCtaRep.frx":0417
         Left            =   3720
         List            =   "frmEnvioEstadoCtaRep.frx":0421
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Agencia:"
         Height          =   255
         Left            =   5640
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Mes: "
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Año:"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Producto:"
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin SICMACT.FlexEdit feCreditos 
      Height          =   3015
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   9495
      _extentx        =   16748
      _extenty        =   5318
      cols0           =   30
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   $"frmEnvioEstadoCtaRep.frx":049D
      encabezadosanchos=   "0-350-2560-2200-3000-950-1000-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200-900-0-0-0-0-0-0-0-0-0-0-0-0"
      font            =   "frmEnvioEstadoCtaRep.frx":05BD
      font            =   "frmEnvioEstadoCtaRep.frx":05E9
      font            =   "frmEnvioEstadoCtaRep.frx":0615
      font            =   "frmEnvioEstadoCtaRep.frx":0641
      font            =   "frmEnvioEstadoCtaRep.frx":066D
      fontfixed       =   "frmEnvioEstadoCtaRep.frx":0699
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      tipobusqueda    =   6
      columnasaeditar =   "X-1-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-17-X-X-X-X-X-X-X-X-X-X-X-X"
      listacontroles  =   "0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-1-0-0-0-0-0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-L-L-C-L-C-C-C-R-R-R-R-R-R-C-C-R-C-C-C-C-C-C-C-C-C-C-C-C-C"
      formatosedit    =   "0-0-0-0-0-0-0-0-2-2-2-2-2-2-0-0-2-0-0-0-0-0-0-0-0-0-0-0-0-0"
      textarray0      =   "dVigencia"
      lbeditarflex    =   -1
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
   Begin SICMACT.FlexEdit feAhorros 
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   9495
      _extentx        =   16748
      _extenty        =   5318
      cols0           =   9
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "Nº-Sel-Cliente-Cuenta-Dirección-Generación-cPersCod-Provincia-Distrito"
      encabezadosanchos=   "400-350-2560-2200-3000-900-0-0-0"
      font            =   "frmEnvioEstadoCtaRep.frx":06C7
      font            =   "frmEnvioEstadoCtaRep.frx":06F3
      font            =   "frmEnvioEstadoCtaRep.frx":071F
      font            =   "frmEnvioEstadoCtaRep.frx":074B
      font            =   "frmEnvioEstadoCtaRep.frx":0777
      fontfixed       =   "frmEnvioEstadoCtaRep.frx":07A3
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      tipobusqueda    =   6
      columnasaeditar =   "X-1-X-X-X-5-X-X-X"
      listacontroles  =   "0-4-0-0-0-1-0-0-0"
      encabezadosalineacion=   "C-C-L-C-L-R-C-C-C"
      formatosedit    =   "0-0-0-0-0-0-0-0-0"
      textarray0      =   "Nº"
      lbeditarflex    =   -1
      colwidth0       =   405
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
End
Attribute VB_Name = "frmEnvioEstadoCtaRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmEnvioEstadoCtaRep
'** Descripción : Formulario para aprobar elegir el tipo de envio de estado de cuenta TI-ERS057-2013
'** Creación : JUEZ, 20130606 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim oEnvEstCta As COMDCaptaGenerales.DCOMCaptaGenerales
Dim rs As ADODB.Recordset
Dim feDatos As FlexEdit
Dim lnTipo As Integer
'COMENTADO POR APRI2018 ERS036-2017
'Private Sub cboAgencia_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        cmdBuscar.SetFocus
'    End If
'End Sub
'
'Private Sub cboPeriodo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        cboRango.SetFocus
'    End If
'End Sub
'
'Private Sub cboRango_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        cboAgencia.SetFocus
'    End If
'End Sub
'
'Private Sub cboTipo_Click()
'    If Trim(Right(Me.cboTipo.Text, 2)) = "2" Then
'        cboPeriodo.Enabled = True
'    Else
'        cboPeriodo.Enabled = False
'        cboPeriodo.ListIndex = -1
'    End If
'End Sub
'
'Private Sub cboTipo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        txtAnioMes.SetFocus
'    End If
'End Sub

Private Sub chkTodos_Click()
    Call HabilitaCheck(feDatos, IIf(chkTodos.value = 1, "1", "0"))
End Sub

Private Sub cmdBuscar_Click()
    Dim lnFila As Integer
    If ValidaDatos Then
        Set oEnvEstCta = New COMDCaptaGenerales.DCOMCaptaGenerales
        'Set rs = oEnvEstCta.RecuperaDatosEnvioEstadoCtaReporte(CInt(Trim(Right(cboTipo.Text, 2))), Replace(txtAnioMes.Text, "/", ""), CInt(IIf(Trim(Right(cboTipo.Text, 2)) = "1", 0, Trim(Right(cboPeriodo.Text, 2)))), CInt(Trim(Right(cboRango.Text, 2))), Trim(Right(cboAgencia.Text, 2)))
        Set rs = oEnvEstCta.RecuperaDatosEnvioEstadoCtaReporte(CInt(Trim(Right(cboTipo.Text, 2))), Trim(txtAnio.Text & IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0" & Trim(Right(cboMes.Text, 2)), Trim(Right(cboMes.Text, 2)))), 0, 2, gsCodAge)
        Set oEnvEstCta = Nothing
        
        Call LimpiaFlex(feAhorros)
        Call LimpiaFlex(feCreditos)
        
        If Not rs.EOF Then
            lnTipo = CInt(Trim(Right(cboTipo.Text, 2)))
            If CInt(Trim(Right(cboTipo.Text, 2))) = 1 Then
                feAhorros.Visible = True
                feCreditos.Visible = False
                Set feDatos = feAhorros
            Else
                feAhorros.Visible = False
                feCreditos.Visible = True
                Set feDatos = feCreditos
            End If
        
            chkTodos.Enabled = True
            'cmdMarcarSelec.Enabled = True 'COMENTADO POR APRI ERS036-2017
        
            Do While Not rs.EOF
                feDatos.AdicionaFila
                lnFila = feDatos.row
                feDatos.TextMatrix(lnFila, 1) = "1"
                feDatos.TextMatrix(lnFila, 2) = rs!cPersNombre
                feDatos.TextMatrix(lnFila, 3) = rs!cCtaCod
                feDatos.TextMatrix(lnFila, 4) = rs!cPersDireccDomicilio
                If lnTipo = 1 Then
                    feDatos.TextMatrix(lnFila, 6) = rs!cPersCod
                    'APRI20180415 ERS036-2017
                    feDatos.TextMatrix(lnFila, 7) = rs!cProvincia
                    feDatos.TextMatrix(lnFila, 8) = rs!cDistrito
                    'END APRI
                End If
                If CInt(Trim(Right(cboTipo.Text, 2))) = 2 Then
                    feDatos.TextMatrix(lnFila, 5) = rs!nCuotasPag
                    feDatos.TextMatrix(lnFila, 6) = rs!nCuotasPend
                    feDatos.TextMatrix(lnFila, 7) = rs!nUltCuotaPag
                    feDatos.TextMatrix(lnFila, 8) = Format(rs!nCapital, "#,##0.00")
                    feDatos.TextMatrix(lnFila, 9) = Format(rs!nIntComp, "#,##0.00")
                    feDatos.TextMatrix(lnFila, 10) = Format(rs!nIntMora, "#,##0.00")
                    feDatos.TextMatrix(lnFila, 11) = Format(rs!nGastos, "#,##0.00")
                    feDatos.TextMatrix(lnFila, 12) = Format(rs!nIntGracia, "#,##0.00")
                    feDatos.TextMatrix(lnFila, 13) = Format(rs!nIntSusp, "#,##0.00")
                    feDatos.TextMatrix(lnFila, 14) = Format(rs!dFecPago, "dd/mm/yyyy")
                    feDatos.TextMatrix(lnFila, 15) = rs!nNroProxCuota
                    feDatos.TextMatrix(lnFila, 16) = Format(rs!nMontoPagado, "#,##0.00")
                    'APRI20180415 ERS036-2017
                    feDatos.TextMatrix(lnFila, 18) = rs!cProvincia
                    feDatos.TextMatrix(lnFila, 19) = rs!cDistrito
                    feDatos.TextMatrix(lnFila, 20) = rs!nMontoCol
                    feDatos.TextMatrix(lnFila, 21) = rs!TEA
                    feDatos.TextMatrix(lnFila, 22) = rs!TCEA
                    feDatos.TextMatrix(lnFila, 23) = rs!nPlazo
                    feDatos.TextMatrix(lnFila, 24) = rs!nComisionEEC
                    feDatos.TextMatrix(lnFila, 25) = rs!nSaldoCap
                    feDatos.TextMatrix(lnFila, 26) = rs!dFechaProxCuota
                    feDatos.TextMatrix(lnFila, 27) = rs!nMontoProxCuota
                    feDatos.TextMatrix(lnFila, 28) = rs!cPersCod
                    feDatos.TextMatrix(lnFila, 29) = rs!dVigencia
                    'END APRI
                End If
                feDatos.TextMatrix(lnFila, IIf(Trim(Right(cboTipo.Text, 2)) = "1", 5, 17)) = "..." 'rs!cPendiente
                rs.MoveNext
            Loop
            feDatos.TopRow = 1
            Frame1.Enabled = False
            chkTodos.Enabled = True
            chkTodos.value = 0
            'cmdMarcarSelec.Enabled = True 'COMENTADO POR APRI ERS036-2017
            Call HabilitaCheck(feDatos, "0")
        Else
            chkTodos.Enabled = False
            'cmdMarcarSelec.Enabled = False 'COMENTADO POR APRI ERS036-2017
            MsgBox "No se encontraron datos", vbInformation, "Aviso" 'APRI ERS036-2017
        End If
    End If
End Sub

Private Sub HabilitaCheck(ByVal pfeDatos As FlexEdit, ByVal psValor As String)
    Dim i As Integer
    For i = 1 To pfeDatos.Rows - 1
        pfeDatos.TextMatrix(i, 1) = psValor
    Next i
End Sub

Private Sub cmdCancelar_Click()
    Frame1.Enabled = True
    chkTodos.value = 0
    Call LimpiaFlex(feAhorros)
    Call LimpiaFlex(feCreditos)
    'APRI ERS036-2017
    txtAnio.Text = ""
    cboMes.ListIndex = -1
    cboTipo.ListIndex = -1
    'End APRI
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGenerar_Click()
If Not (feDatos Is Nothing) Then
    If feDatos.TextMatrix(1, 0) <> "" Then
    
    If MsgBox("¿Está seguro de Generar los Formatos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Dim lsModeloPlantilla As String
    Dim i As Integer
    Dim nCantCheck As Integer
    Dim oMov As COMNContabilidad.NCOMContFunciones
    Dim psMovNroReg As String
    
        For i = 1 To feDatos.Rows - 1
            If feDatos.TextMatrix(i, 1) = "." Then
                nCantCheck = nCantCheck + 1
                'Exit For
            End If
        Next i
        
        If nCantCheck > 0 Then
            i = 0
            'pgb.value = 0
            'pgb.Min = 0
            If lnTipo = 1 Then
                'RecuperaDatosFormatoEstadoCtaAhorros()
    '            Dim xlAplicacion As Excel.Application
    '            Dim xlLibro As Excel.Workbook
    '            Dim lbLibroOpen As Boolean
    '            Dim lsArchivo As String
    '            Dim lsHoja As String
    '            Dim xlHoja1 As Excel.Worksheet
    '            Dim xlHoja2 As Excel.Worksheet
    '            Dim nLin As Long
    '            Dim nItem As Long
    '            Dim sColumna As String
    '
    '            lsArchivo = App.path & "\SPOOLER\EnvioEstadoCtaCap_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    '            lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    '            If Not lbLibroOpen Then
    '                Exit Sub
    '            End If
    '            nLin = 1
    '
    '            lsHoja = "Ahorros"
    '            sColumna = "F"
    '
    '            ExcelAddHoja lsHoja, xlLibro, xlHoja1
    '
    '            xlHoja1.Range("A1:Y1").EntireColumn.Font.FontStyle = "Arial"
    '            xlHoja1.PageSetup.Orientation = xlLandscape
    '            xlHoja1.PageSetup.CenterHorizontally = True
    '            xlHoja1.PageSetup.Zoom = 75
    '            xlHoja1.PageSetup.TopMargin = 2
    '
    '            xlHoja1.Range("A1:A1").RowHeight = 18
    '            xlHoja1.Range("A1:A1").ColumnWidth = 5
    '            xlHoja1.Range("B1:B1").ColumnWidth = 5
    '            xlHoja1.Range("C1:C1").ColumnWidth = 40
    '            xlHoja1.Range("D1:D1").ColumnWidth = 20
    '            xlHoja1.Range("E1:E1").ColumnWidth = 40
    '
    '            For i = 1 To feDatos.Rows - 1
    '                If feDatos.TextMatrix(i, 1) = "." Then
    '                    xlHoja1.Cells(nLin, 1) = "REPORTE ENVIO ESTADO DE CUENTAS " & IIf(lnTipo = 1, "AHORROS", "CREDITOS")
    '                    xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).Merge True
    '                    xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).Font.Bold = True
    '                    xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).HorizontalAlignment = xlHAlignCenter
    '
    '                    nLin = nLin + 2
    '
    '                    xlHoja1.Cells(nLin, 1) = "ITEM"
    '                    xlHoja1.Cells(nLin, 2) = "SEL"
    '                    xlHoja1.Cells(nLin, 3) = "CLIENTE"
    '                    xlHoja1.Cells(nLin, 4) = "CUENTA"
    '                    xlHoja1.Cells(nLin, 5) = "DIRECCION"
    '                    If lnTipo = 2 Then
    '                        xlHoja1.Cells(nLin, 6) = "CUOTAS PAG"
    '                        xlHoja1.Cells(nLin, 7) = "CUOTAS PEND"
    '                        xlHoja1.Cells(nLin, 8) = "ULTI CUOTA PAG"
    '                        xlHoja1.Cells(nLin, 9) = "CAPITAL"
    '                        xlHoja1.Cells(nLin, 10) = "INT COMP"
    '                        xlHoja1.Cells(nLin, 11) = "INT MORA"
    '                        xlHoja1.Cells(nLin, 12) = "GASTOS"
    '                        xlHoja1.Cells(nLin, 13) = "INT GRACIA"
    '                        xlHoja1.Cells(nLin, 14) = "INT SUSP"
    '                        xlHoja1.Cells(nLin, 15) = "FEC PAGO"
    '                        xlHoja1.Cells(nLin, 16) = "PROX CUOTA"
    '                        xlHoja1.Cells(nLin, 17) = "MONTO PAG"
    '                    End If
    '                    xlHoja1.Cells(nLin, IIf(lnTipo = 1, 6, 18)) = "PEND"
    '
    '                    xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).Font.Bold = True
    '                    xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).HorizontalAlignment = xlHAlignCenter
    '                    xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    '                    xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
    '                    xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).Borders(xlInsideVertical).Color = vbBlack
    '                    xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).Interior.Color = RGB(255, 50, 50)
    '                    xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).Font.Color = RGB(255, 255, 255)
    '
    '                    With xlHoja1.PageSetup
    '                        .LeftHeader = ""
    '                        .CenterHeader = ""
    '                        .RightHeader = ""
    '                        .LeftFooter = ""
    '                        .CenterFooter = ""
    '                        .RightFooter = ""
    '
    '                        .PrintHeadings = False
    '                        .PrintGridlines = False
    '                        .PrintComments = xlPrintNoComments
    '                        .CenterHorizontally = True
    '                        .CenterVertically = False
    '                        .Orientation = xlLandscape
    '                        .Draft = False
    '                        .FirstPageNumber = xlAutomatic
    '                        .Order = xlDownThenOver
    '                        .BlackAndWhite = False
    '                        .Zoom = 55
    '                    End With
    '
    '                    nItem = 1
    '                    nLin = nLin + 1
    '                    For nItem = 1 To feDatos.Rows - 1
    '                        xlHoja1.Range("A" & nLin & ":E" & nLin).HorizontalAlignment = xlHAlignLeft
    '                        xlHoja1.Range("I" & nLin & ":N" & nLin).NumberFormat = "#,##0.00"
    '                        xlHoja1.Range("Q" & nLin & ":Q" & nLin).NumberFormat = "#,##0.00"
    '                        xlHoja1.Cells(nLin, 1) = feDatos.TextMatrix(nItem, 0)
    '                        xlHoja1.Cells(nLin, 2) = IIf(feDatos.TextMatrix(nItem, 1) = "", "NO", "SI")
    '                        xlHoja1.Cells(nLin, 3) = feDatos.TextMatrix(nItem, 2)
    '                        xlHoja1.Cells(nLin, 4) = "'" & feDatos.TextMatrix(nItem, 3)
    '                        xlHoja1.Cells(nLin, 5) = feDatos.TextMatrix(nItem, 4)
    '                        If lnTipo = 2 Then
    '                            xlHoja1.Cells(nLin, 6) = feDatos.TextMatrix(nItem, 5)
    '                            xlHoja1.Cells(nLin, 7) = feDatos.TextMatrix(nItem, 6)
    '                            xlHoja1.Cells(nLin, 8) = feDatos.TextMatrix(nItem, 7)
    '                            xlHoja1.Cells(nLin, 9) = Format(feDatos.TextMatrix(nItem, 8), "#,##0.00")
    '                            xlHoja1.Cells(nLin, 10) = Format(feDatos.TextMatrix(nItem, 9), "#,##0.00")
    '                            xlHoja1.Cells(nLin, 11) = Format(feDatos.TextMatrix(nItem, 10), "#,##0.00")
    '                            xlHoja1.Cells(nLin, 12) = Format(feDatos.TextMatrix(nItem, 11), "#,##0.00")
    '                            xlHoja1.Cells(nLin, 13) = Format(feDatos.TextMatrix(nItem, 12), "#,##0.00")
    '                            xlHoja1.Cells(nLin, 14) = Format(feDatos.TextMatrix(nItem, 13), "#,##0.00")
    '                            xlHoja1.Cells(nLin, 15) = feDatos.TextMatrix(nItem, 14)
    '                            xlHoja1.Cells(nLin, 16) = Format(feDatos.TextMatrix(nItem, 15), "#,##0.00")
    '                            xlHoja1.Cells(nLin, 17) = feDatos.TextMatrix(nItem, 16)
    '                        End If
    '                        xlHoja1.Cells(nLin, IIf(lnTipo = 1, 6, 18)) = feDatos.TextMatrix(nItem, IIf(lnTipo = 1, 5, 17))
    '                        nLin = nLin + 1
    '                    Next nItem
    '                End If
    '            Next i
    '
    '            ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
    '            CargaArchivo lsArchivo, App.path & "\SPOOLER\"
            Dim sPeriodo As String
            sPeriodo = Trim(txtAnio.Text & IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0" & Trim(Right(cboMes.Text, 2)), Trim(Right(cboMes.Text, 2))))
             GeneraEstadoCuentaAhorros feDatos, sPeriodo
            For i = 1 To feDatos.Rows - 1
                If feDatos.TextMatrix(i, 1) = "." Then
                    Set oMov = New COMNContabilidad.NCOMContFunciones
                    psMovNroReg = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                    Set oEnvEstCta = New COMDCaptaGenerales.DCOMCaptaGenerales
                    Call oEnvEstCta.InsertarGeneracionEnvioEstadoCta(feDatos.TextMatrix(i, 3), Trim(feDatos.TextMatrix(i, 6)), Trim(txtAnio.Text), Trim(IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0" & Trim(Right(cboMes.Text, 2)), Trim(Right(cboMes.Text, 2)))), 0, psMovNroReg)
                    Set oEnvEstCta = Nothing
                End If
            Next i
             
'**********************COMENTADO POR APRI ERS036-2017***************************************
'                Dim fs As Scripting.FileSystemObject
'                Dim xlsAplicacion As Excel.Application
'                Dim lsArchivo As String
'                Dim lsFile As String
'                Dim lsNomHoja As String
'                Dim xlsLibro As Excel.Workbook
'                Dim xlHoja1 As Excel.Worksheet
'                Dim lbExisteHoja As Boolean
'                Dim lnFila As Long
'                Dim cCabFechas As String
'                Dim nCantRegistros As Integer
'                Dim nItemCta As Integer
'                Dim nFilasReg As Integer
'
'                Set fs = New Scripting.FileSystemObject
'                Set xlsAplicacion = New Excel.Application
'
'                lsNomHoja = "Hoja1"
'                lsFile = "EnvioEstadoCtaCap"
'
'                lsArchivo = "\spooler\" & "EnvioEstadoCtaCap_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
'                If fs.FileExists(App.Path & "\FormatoCarta\" & lsFile & ".xls") Then
'                    Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsFile & ".xls")
'                Else
'                    MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
'                    Exit Sub
'                End If
'
'                For Each xlHoja1 In xlsLibro.Worksheets
'                   If xlHoja1.Name = lsNomHoja Then
'                        xlHoja1.Activate
'                     lbExisteHoja = True
'                    Exit For
'                   End If
'                Next
'
'                If lbExisteHoja = False Then
'                    Set xlHoja1 = xlsLibro.Worksheets
'                    xlHoja1.Name = lsNomHoja
'                End If
'
'    '            xlHoja1.Range("A1:H1").EntireColumn.Font.FontStyle = "Tahoma"
'    '            xlHoja1.Range("A1:H1").EntireColumn.Font.FontStyle = "Tahoma"
'    '            xlHoja1.PageSetup.Orientation = xlLandscape
'    '            xlHoja1.PageSetup.CenterHorizontally = True
'    '            xlHoja1.PageSetup.Zoom = 75
'    '            xlHoja1.PageSetup.TopMargin = 2
'                lnFila = 1
'                nCantRegistros = 36
'                nFilasReg = 11
'                pgb.Max = nCantCheck
'                For i = 1 To feDatos.Rows - 1
'                    If feDatos.TextMatrix(i, 1) = "." Then
'
'                        Set oEnvEstCta = New COMDCaptaGenerales.DCOMCaptaGenerales
'                        Set rs = oEnvEstCta.RecuperaDatosFormatoEstadoCta(feDatos.TextMatrix(i, 3), 1, Replace(txtAnioMes.Text, "/", ""), feDatos.TextMatrix(i, 6))
'                        Set oEnvEstCta = Nothing
'                        nItemCta = nItemCta + 1
'                        If nItemCta > 1 Then
'                            cCabFechas = xlHoja1.Cells(2, 1)
'                            xlHoja1.Cells(2, 1) = ""
'                            xlHoja1.Range("A3:H3").Interior.Color = RGB(255, 255, 255)
'                            xlHoja1.Range("A1", "H2").CopyPicture
'                            xlHoja1.Cells(2, 1) = cCabFechas
'                            xlHoja1.Range("A3:H3").Interior.Color = RGB(255, 0, 0)
'                            xlHoja1.Range("A" & lnFila, "A" & lnFila).RowHeight = 26.5
'                            xlHoja1.Range("A" & lnFila, "A" & lnFila + 1).PasteSpecial
'                            nCantRegistros = nCantRegistros
'                            'lnFila = lnFila + 1
'                        End If
'
'                        lnFila = lnFila + 1
'
'                        xlHoja1.Cells(lnFila, 1) = "Del " & Format(rs!dFecIni, "dd/mm/yyyy") & " al " & Format(rs!dFecFin, "dd/mm/yyyy")
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Size = 9
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).VerticalAlignment = xlVAlignTop
'
'                        lnFila = lnFila + 1
'                        xlHoja1.Cells(lnFila, 1) = "INFORMACION GENERAL"
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).VerticalAlignment = xlVAlignCenter
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).RowHeight = 26.25
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Merge True
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Bold = True
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Interior.Color = RGB(255, 0, 0)
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Color = RGB(255, 255, 255)
'                        lnFila = lnFila + 2
'
'                        xlHoja1.Cells(lnFila, 1) = "Cliente:"
'                        xlHoja1.Cells(lnFila, 3) = Trim(rs!cPersNombre)
'                        xlHoja1.Range("C" & lnFila & ":D" & lnFila).Merge True
'                        xlHoja1.Range("C" & lnFila & ":D" & lnFila).Borders(xlEdgeBottom).LineStyle = xlDash
'                        xlHoja1.Range("C" & lnFila & ":D" & lnFila).Borders(xlEdgeRight).LineStyle = xlDash
'                        xlHoja1.Range("C" & lnFila & ":D" & lnFila).Borders(xlEdgeBottom).Color = vbRed
'                        xlHoja1.Range("C" & lnFila & ":D" & lnFila).Borders(xlEdgeRight).Color = vbRed
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Size = 9
'                        xlHoja1.Cells(lnFila, 5) = "Cuenta:"
'                        xlHoja1.Cells(lnFila, 6) = Trim("'" & rs!cCtaCod)
'                        xlHoja1.Range("E" & lnFila & ":E" & lnFila).HorizontalAlignment = xlHAlignRight
'                        xlHoja1.Range("F" & lnFila & ":H" & lnFila).Merge True
'                        xlHoja1.Range("F" & lnFila & ":H" & lnFila).Borders(xlEdgeBottom).LineStyle = xlDash
'                        xlHoja1.Range("F" & lnFila & ":H" & lnFila).Borders(xlEdgeRight).LineStyle = xlDash
'                        xlHoja1.Range("F" & lnFila & ":H" & lnFila).Borders(xlEdgeBottom).Color = vbRed
'                        xlHoja1.Range("F" & lnFila & ":H" & lnFila).Borders(xlEdgeRight).Color = vbRed
'                        lnFila = lnFila + 2
'
'                        xlHoja1.Cells(lnFila, 1) = "Dirección:"
'                        xlHoja1.Cells(lnFila, 3) = Trim(rs!cPersDireccDomicilio)
'                        xlHoja1.Range("C" & lnFila & ":D" & lnFila).Merge True
'                        xlHoja1.Range("C" & lnFila & ":D" & lnFila).Borders(xlEdgeBottom).LineStyle = xlDash
'                        xlHoja1.Range("C" & lnFila & ":D" & lnFila).Borders(xlEdgeRight).LineStyle = xlDash
'                        xlHoja1.Range("C" & lnFila & ":D" & lnFila).Borders(xlEdgeBottom).Color = vbRed
'                        xlHoja1.Range("C" & lnFila & ":D" & lnFila).Borders(xlEdgeRight).Color = vbRed
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Size = 9
'                        xlHoja1.Cells(lnFila, 5) = "Periodo:"
'                        xlHoja1.Cells(lnFila, 6) = Choose(Right(rs!cPeriodo, 2), "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre") & " - " & Left(rs!cPeriodo, 4)
'                        xlHoja1.Range("E" & lnFila & ":E" & lnFila).HorizontalAlignment = xlHAlignRight
'                        xlHoja1.Range("F" & lnFila & ":H" & lnFila).Merge True
'                        xlHoja1.Range("F" & lnFila & ":H" & lnFila).Borders(xlEdgeBottom).LineStyle = xlDash
'                        xlHoja1.Range("F" & lnFila & ":H" & lnFila).Borders(xlEdgeRight).LineStyle = xlDash
'                        xlHoja1.Range("F" & lnFila & ":H" & lnFila).Borders(xlEdgeBottom).Color = vbRed
'                        xlHoja1.Range("F" & lnFila & ":H" & lnFila).Borders(xlEdgeRight).Color = vbRed
'                        lnFila = lnFila + 2
'
'                        xlHoja1.Cells(lnFila, 1) = "Movimientos de la Cuenta"
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).VerticalAlignment = xlVAlignCenter
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).RowHeight = 28.25
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Merge True
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Bold = True
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Interior.Color = RGB(255, 0, 0)
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Color = RGB(255, 255, 255)
'                        lnFila = lnFila + 2
'
'                        xlHoja1.Cells(lnFila, 1) = "Nº"
'                        xlHoja1.Cells(lnFila, 2) = "Fechas"
'                        xlHoja1.Cells(lnFila, 3) = "Tipo de operación"
'                        xlHoja1.Cells(lnFila, 4) = "Depósitos"
'                        xlHoja1.Cells(lnFila, 5) = "Retiros"
'                        xlHoja1.Cells(lnFila, 6) = "Saldo Cont"
'                        xlHoja1.Cells(lnFila, 7) = "Agencia"
'                        xlHoja1.Cells(lnFila, 8) = "Usuario"
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Interior.Color = RGB(255, 0, 0)
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Color = RGB(255, 255, 255)
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Size = 8
'
'                        Dim rsMov As ADODB.Recordset
'                        Set oEnvEstCta = New COMDCaptaGenerales.DCOMCaptaGenerales
'                        Set rsMov = oEnvEstCta.RecuperaDatosFormatoEstadoCtaAhorrosMov(feDatos.TextMatrix(i, 3), Replace(txtAnioMes.Text, "/", ""))
'                        Set oEnvEstCta = Nothing
'
'                        Do While Not rsMov.EOF
'                            lnFila = lnFila + 1
'                            xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Size = 8
'                            xlHoja1.Range("D" & lnFila & ":F" & lnFila).NumberFormat = "#,##0.00"
'                            xlHoja1.Cells(lnFila, 1) = rsMov!nItem
'                            xlHoja1.Cells(lnFila, 2) = Format(rsMov!Fecha, "dd/mm/yyyy")
'                            xlHoja1.Cells(lnFila, 3) = Trim(rsMov!Operacion)
'                            xlHoja1.Cells(lnFila, 4) = Format(rsMov!nDep, "#,##0.00")
'                            xlHoja1.Cells(lnFila, 5) = Format(rsMov!nRet, "#,##0.00")
'                            xlHoja1.Cells(lnFila, 6) = Format(rsMov!nSaldoContable, "#,##0.00")
'                            xlHoja1.Cells(lnFila, 7) = rsMov!CAgencia
'                            xlHoja1.Cells(lnFila, 8) = rsMov!cUsu
'                            rsMov.MoveNext
'                        Loop
'                        If rsMov.RecordCount > 0 Then
'                            rsMov.MoveFirst
'                        End If
'
'                        lnFila = IIf(lnFila < (nCantRegistros * nItemCta) + IIf(nItemCta > 1, nFilasReg, 0), (nCantRegistros * nItemCta) + IIf(nItemCta > 1, nFilasReg, 0), lnFila)
'
'                        lnFila = lnFila + 2
'                        xlHoja1.Cells(lnFila, 1) = "TOTAL"
'                        xlHoja1.Cells(lnFila, 6) = Format(IIf(rsMov.RecordCount > 0, rsMov!nTotal, 0), "#,##0.00")
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Size = 9
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Bold = True
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).RowHeight = 19.5
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).HorizontalAlignment = xlHAlignRight
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).Borders(xlEdgeBottom).LineStyle = xlDash
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).Borders(xlEdgeBottom).Color = vbRed
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).NumberFormat = "#,##0.00"
'                        lnFila = lnFila + 1
'                        xlHoja1.Cells(lnFila, 1) = "SALDO DISPONIBLE"
'                        xlHoja1.Cells(lnFila, 6) = Format(rs!nSaldoDisp, "#,##0.00")
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Size = 9
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).RowHeight = 19.5
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).HorizontalAlignment = xlHAlignRight
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).Borders(xlEdgeBottom).LineStyle = xlDash
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).Borders(xlEdgeBottom).Color = vbRed
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).NumberFormat = "#,##0.00"
'                        lnFila = lnFila + 1
'                        xlHoja1.Cells(lnFila, 1) = "SALDO CONTABLE"
'                        xlHoja1.Cells(lnFila, 6) = Format(rs!nSaldoContable, "#,##0.00")
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Size = 9
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).RowHeight = 19.5
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).HorizontalAlignment = xlHAlignRight
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).Borders(xlEdgeBottom).LineStyle = xlDash
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).Borders(xlEdgeBottom).Color = vbRed
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).NumberFormat = "#,##0.00"
'                        lnFila = lnFila + 1
'                        xlHoja1.Cells(lnFila, 1) = "BLOQUEO PARCIAL"
'                        xlHoja1.Cells(lnFila, 6) = Format(rs!nBloqueoParcial, "#,##0.00")
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Size = 9
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).RowHeight = 19.5
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).HorizontalAlignment = xlHAlignRight
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).Borders(xlEdgeBottom).LineStyle = xlDash
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).Borders(xlEdgeBottom).Color = vbRed
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).NumberFormat = "#,##0.00"
'                        lnFila = lnFila + 1
'                        xlHoja1.Cells(lnFila, 1) = "INTERES DEL MES"
'                        xlHoja1.Cells(lnFila, 6) = Format(IIf(rsMov.RecordCount > 0, rsMov!nInteresGan, 0), "#,##0.00")
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Size = 9
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).RowHeight = 19.5
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).HorizontalAlignment = xlHAlignRight
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).Borders(xlEdgeBottom).LineStyle = xlDash
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).Borders(xlEdgeBottom).Color = vbRed
'                        xlHoja1.Range("F" & lnFila & ":F" & lnFila).NumberFormat = "#,##0.00"
'                        lnFila = lnFila + 1
'
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Borders(xlEdgeBottom).LineStyle = xlContinuous
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Borders(xlEdgeBottom).Color = RGB(166, 166, 166)
'                        lnFila = lnFila + 1
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).RowHeight = 10.5
'                        lnFila = lnFila + 1
'
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).RowHeight = 83.25
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Font.Size = 9
'                        xlHoja1.Cells(lnFila, 1) = "Advertencia: Si dentro de 30 días no se formula observación al presente estado, daremos por conforme la cuenta "
'                        xlHoja1.Cells(lnFila, 1) = xlHoja1.Cells(lnFila, 1) + "y aprobado el saldo. En caso contrario sirvase dirigirse a nuestras oficinas para atender sus observaciones." & Chr(10) & Chr(10)
'                        'xlHoja1.Cells(lnFila, 1) = xlHoja1.Cells(lnFila, 1) + "                                                                                                            "
'                        xlHoja1.Cells(lnFila, 1) = xlHoja1.Cells(lnFila, 1) + "En caso de Reclamos y Servicios , el cliente podrá recurrir indistintamente a las siguientes instancias :" & Chr(10) & Chr(10)
'    '                    xlHoja1.Cells(lnFila, 1) = xlHoja1.Cells(lnFila, 1) + "                                                                                                            "
'                        xlHoja1.Cells(lnFila, 1) = xlHoja1.Cells(lnFila, 1) + "1. A nuestra red de oficinas; 2. INDECOPI; 3. A la plataforma de Atención al Usuario de la Superintendencia de banca, "
'                        xlHoja1.Cells(lnFila, 1) = xlHoja1.Cells(lnFila, 1) + "Seguros y AFP."
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Merge True
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).HorizontalAlignment = xlHAlignLeft
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).VerticalAlignment = xlVAlignTop
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Borders(xlEdgeBottom).LineStyle = xlContinuous
'                        xlHoja1.Range("A" & lnFila & ":H" & lnFila).Borders(xlEdgeBottom).Color = RGB(166, 166, 166)
'                        lnFila = lnFila + 3
'                        'If nItemCta Mod 2 = 0 Then
'                        If nItemCta > 1 Then
'                            nFilasReg = nFilasReg + 11
'                        End If
'
'                        Dim loContFunct As COMNContabilidad.NCOMContFunciones
'                        Dim psMovNro As String
'                        Set loContFunct = New COMNContabilidad.NCOMContFunciones
'                        psMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'                        Set oEnvEstCta = New COMDCaptaGenerales.DCOMCaptaGenerales
'                        Call oEnvEstCta.InsertarGeneracionEnvioEstadoCta(feDatos.TextMatrix(i, 3), feDatos.TextMatrix(i, 6), Left(txtAnioMes.Text, 4), Right(txtAnioMes.Text, 2), 0, psMovNro)
'                        Set oEnvEstCta = Nothing
'                        pgb.value = pgb.value + 1
'                    End If
'                Next i
'
'                Dim psArchivoAGrabarC As String
'
'                xlHoja1.SaveAs App.Path & lsArchivo
'                psArchivoAGrabarC = App.Path & lsArchivo
'                xlsAplicacion.Visible = True
'                xlsAplicacion.Windows(1).Visible = True
'                Set xlsAplicacion = Nothing
'                Set xlsLibro = Nothing
'                Set xlHoja1 = Nothing
                MsgBox "Se han generado los formatos con exito", vbInformation, "Aviso"
                'pgb.value = 0
                'pgb.Min = 0
                cmdCancelar_Click
            Else
            GeneraEstadoCuentaCredito feDatos
            
            For i = 1 To feDatos.Rows - 1
                If feDatos.TextMatrix(i, 1) = "." Then
                    
                    Set oMov = New COMNContabilidad.NCOMContFunciones
                    psMovNroReg = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                    Set oEnvEstCta = New COMDCaptaGenerales.DCOMCaptaGenerales
                    Call oEnvEstCta.InsertarGeneracionEnvioEstadoCta(feDatos.TextMatrix(i, 3), Trim(feDatos.TextMatrix(i, 28)), Trim(txtAnio.Text), Trim(IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0" & Trim(Right(cboMes.Text, 2)), Trim(Right(cboMes.Text, 2)))), 0, psMovNroReg)
                    Set oEnvEstCta = Nothing
                End If
            Next i
                       
            
'**********************COMENTADO POR APRI ERS036-2017***************************************
'                lsModeloPlantilla = App.Path & "\FormatoCarta\EnvioEstadoCtaCred.doc"
'
'                'Crea una clase que de Word Object
'                Dim wApp As Word.Application
'                Dim wAppSource As Word.Application
'                'Create a new instance of word
'                Set wApp = New Word.Application
'                Set wAppSource = New Word.Application
'
'                Dim RangeSource As Word.Range
'                'Abre Documento Plantilla
'                wAppSource.Documents.Open FileName:=lsModeloPlantilla
'                Set RangeSource = wAppSource.ActiveDocument.Content
'                'Lo carga en Memoria
'                wAppSource.ActiveDocument.Content.Copy
'
'                'Crea Nuevo Documento
'                wApp.Documents.Add
'
'                With wApp.ActiveDocument.PageSetup
'                    .LeftMargin = CentimetersToPoints(1.5)
'                    .RightMargin = CentimetersToPoints(1)
'                    .TopMargin = CentimetersToPoints(1.5)
'                    .BottomMargin = CentimetersToPoints(1)
'                End With
'
'                pgb.Max = nCantCheck
'                For i = 1 To feDatos.Rows - 1
'                    If feDatos.TextMatrix(i, 1) = "." Then
'                        wApp.Application.Selection.TypeParagraph
'                        wApp.Application.Selection.PasteAndFormat (wdPasteDefault)
'                        wApp.Application.Selection.InsertBreak
'                        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
'                        wApp.Selection.MoveEnd
'
'                        Set oEnvEstCta = New COMDCaptaGenerales.DCOMCaptaGenerales
'                        Set rs = oEnvEstCta.RecuperaDatosFormatoEstadoCta(feDatos.TextMatrix(i, 3), 2)
'                        Set oEnvEstCta = Nothing
'
'                        'Monto Prestamo
'                        With wApp.Selection.Find
'                            .Text = "<<nMontoPrestamo>>"
'                            .Replacement.Text = IIf(Mid(rs!cCtaCod, 9, 1) = "1", "S/.", "$") & " " & Format(rs!nMontoCol, "#,##0.00")
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Cliente
'                        With wApp.Selection.Find
'                            .Text = "<<cPersNombre>>"
'                            .Replacement.Text = Trim(rs!cPersNombre)
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Cuenta
'                        With wApp.Selection.Find
'                            .Text = "<<cCtaCod>>"
'                            .Replacement.Text = Trim(rs!cCtaCod)
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Direccion
'                        With wApp.Selection.Find
'                            .Text = "<<cPersDireccDomicilio>>"
'                            .Replacement.Text = Trim(rs!cPersDireccDomicilio)
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Periodo
'                        With wApp.Selection.Find
'                            .Text = "<<cPeriodo>>"
'                            .Replacement.Text = Choose(Right(rs!cPeriodo, 2), "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre") & " - " & Left(rs!cPeriodo, 4)
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Cuotas Pagadas
'                        With wApp.Selection.Find
'                            .Text = "<<nCuotasPag>>"
'                            .Replacement.Text = rs!nCuotasPag
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Cuotas Pendientes
'                        With wApp.Selection.Find
'                            .Text = "<<nCuotasPend>>"
'                            .Replacement.Text = rs!nCuotasPend
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Cuota Pagada
'                        With wApp.Selection.Find
'                            .Text = "<<nCuota>>"
'                            .Replacement.Text = Format(rs!nCuotaPag, "#,##0.00")
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Capital
'                        With wApp.Selection.Find
'                            .Text = "<<nCapital>>"
'                            .Replacement.Text = Format(rs!nCapital, "#,##0.00")
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Int Comp
'                        With wApp.Selection.Find
'                            .Text = "<<nIntComp>>"
'                            .Replacement.Text = Format(rs!nIntComp, "#,##0.00")
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Int Mora
'                        With wApp.Selection.Find
'                            .Text = "<<nIntMora>>"
'                            .Replacement.Text = Format(rs!nIntMora, "#,##0.00")
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Gastos
'                        With wApp.Selection.Find
'                            .Text = "<<nGastos>>"
'                            .Replacement.Text = Format(rs!nGastos, "#,##0.00")
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Int Gracia
'                        With wApp.Selection.Find
'                            .Text = "<<nIntGracia>>"
'                            .Replacement.Text = Format(rs!nIntGracia, "#,##0.00")
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Int Susp y Reprog
'                        With wApp.Selection.Find
'                            .Text = "<<nIntSusp>>"
'                            .Replacement.Text = Format(rs!nIntSusp, "#,##0.00")
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Comision
'                        With wApp.Selection.Find
'                            .Text = "<<nComision>>"
'                            .Replacement.Text = Format(rs!nComision, "#,##0.00")
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Total Pagado
'                        With wApp.Selection.Find
'                            .Text = "<<nTotalPag>>"
'                            .Replacement.Text = Format(rs!nMontoPagado, "#,##0.00")
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Fecha Pago
'                        With wApp.Selection.Find
'                            .Text = "<<dFecPago>>"
'                            .Replacement.Text = IIf(Format(rs!dFecPago, "dd/mm/yyyy") = "01/01/1900", "", Format(rs!dFecPago, "dd/mm/yyyy"))
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Saldo Capital
'                        With wApp.Selection.Find
'                            .Text = "<<nSaldo>>"
'                            .Replacement.Text = Format(rs!nSaldoCap, "#,##0.00")
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Fecha Couta Pendiente
'                        With wApp.Selection.Find
'                            .Text = "<<dFecha>>"
'                            .Replacement.Text = Format(rs!dVenc, "yyyy-mm-dd")
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Descripcion Cuota Pendiente
'                        With wApp.Selection.Find
'                            .Text = "<<cCtaDesc>>"
'                            .Replacement.Text = Trim(rs!cCtaDesc)
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        'Importe Proxima Cuota
'                        With wApp.Selection.Find
'                            .Text = "<<nImporte>>"
'                            .Replacement.Text = Format(rs!nImporte, "#,##0.00")
'                            .Forward = True
'                            .Wrap = wdFindContinue
'                            .Format = False
'                            .Execute Replace:=wdReplaceAll
'                        End With
'
'                        Dim oMov As COMNContabilidad.NCOMContFunciones
'                        Dim psMovNroReg As String
'                        Set oMov = New COMNContabilidad.NCOMContFunciones
'                        psMovNroReg = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'                        Set oEnvEstCta = New COMDCaptaGenerales.DCOMCaptaGenerales
'                        Call oEnvEstCta.InsertarGeneracionEnvioEstadoCta(feDatos.TextMatrix(i, 3), rs!cPersCod, Left(txtAnioMes.Text, 4), Right(txtAnioMes.Text, 2), CInt(Trim(Right(cboPeriodo.Text, 2))), psMovNroReg)
'                        Set oEnvEstCta = Nothing
'                    End If
'                    pgb.value = pgb.value + 1
'                Next i
'
'                wAppSource.ActiveDocument.Close
'                wAppSource.Quit
'                wApp.ActiveDocument.SaveAs (App.Path & "\SPOOLER\EnvioEstadoCtaCred_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".doc")
'                wApp.Visible = True
'
'                Set wAppSource = Nothing
'                Set wApp = Nothing
                MsgBox "Se han generado los formatos con exito", vbInformation, "Aviso"
                'pgb.value = 0
                'pgb.Min = 0
                cmdCancelar_Click
            End If
        Else
            MsgBox "No seleccionó ningún registro de la lista", vbInformation, "Aviso"
        End If
    Else
        MsgBox "No existen datos para generar formatos", vbInformation, "Aviso"
    End If
Else
    MsgBox "No existen datos para generar formatos", vbInformation, "Aviso"
End If
End Sub

'APRI ERS036-2017
Private Sub feAhorros_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    frmEnvioEstadoCtaRepHistorial.Inicio (CStr(feAhorros.TextMatrix(feAhorros.row, 3)))
End Sub
Private Sub feCreditos_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    frmEnvioEstadoCtaRepHistorial.Inicio (CStr(feCreditos.TextMatrix(feCreditos.row, 3)))
End Sub

Private Sub Form_Load()
    Dim oAge As COMDConstantes.DCOMAgencias

'COMENTADO POR APRI 20180415 ERS026-2017
'    Set oAge = New COMDConstantes.DCOMAgencias
'    Set rs = oAge.ObtieneAgencias()
'    Call Llenar_Combo_Agencia_con_Recordset(rs, cboAgencia)
'    Set oAge = Nothing
'
'    Dim oCons As COMDConstantes.DCOMConstantes
'    Set oCons = New COMDConstantes.DCOMConstantes
'    Set rs = oCons.RecuperaConstantes(9111)
'    Call Llenar_Combo_con_Recordset(rs, cboPeriodo)
'    Set rs = oCons.RecuperaConstantes(9112)
'    Call Llenar_Combo_con_Recordset(rs, cboRango)

    'APRI20180415 ERS036-2017
    Dim oCons As COMDConstantes.DCOMConstantes
    Set oCons = New COMDConstantes.DCOMConstantes
    Set rs = oCons.RecuperaConstantes(1010)
    Call Llenar_Combo_con_Recordset(rs, cboMes)
    txtAgencia.Text = UCase(gsNomAge)
    txtAgencia.Enabled = False
    
    Set feDatos = Nothing
    feAhorros.Visible = True
    feCreditos.Visible = False
    
    'cboTipo.ListIndex = 0
    chkTodos.Enabled = False
    chkTodos.value = 0
    'cmdMarcarSelec.Enabled = False
    lnTipo = 0
End Sub

'COMENTADO POR APRI20180415 ERS026-2017
'Private Sub txtAnioMes_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If CInt(Trim(Right(Me.cboTipo.Text, 2))) = 1 Then
'            cboRango.SetFocus
'        Else
'            cboPeriodo.SetFocus
'        End If
'    End If
'End Sub
'END APRI

Public Function ValidaDatos() As Boolean
    ValidaDatos = False
    
    If cboMes.Text = "" Then
        MsgBox "Debe elegir el mes", vbInformation, "Aviso"
        cboMes.SetFocus
        Exit Function
    End If
    If txtAnio.Text = "" Then
        MsgBox "Debe ingresar el año", vbInformation, "Aviso"
        txtAnio.SetFocus
        Exit Function
    End If
    If cboTipo.Text = "" Then
        'MsgBox "Debe elegir el tipo", vbInformation, "Aviso"
        MsgBox "Debe elegir el tipo de producto", vbInformation, "Aviso"
        cboTipo.SetFocus
        Exit Function
    End If
    
'COMENTADO POR APRI20180415 ERS026-2017
'    If txtAnioMes.Text = "____/__" Then
'        MsgBox "Debe ingresar el año y el mes", vbInformation, "Aviso"
'        txtAnioMes.SetFocus
'        Exit Function
'    End If
'    If Trim(Right(cboTipo.Text, 2)) = "2" Then
'        If cboPeriodo.Text = "" Then
'            MsgBox "Debe elegir el periodo", vbInformation, "Aviso"
'            cboPeriodo.SetFocus
'            Exit Function
'        End If
'    End If
'    If cboRango.Text = "" Then
'        MsgBox "Debe elegir el rango", vbInformation, "Aviso"
'        cboRango.SetFocus
'        Exit Function
'    End If
'    If cboAgencia.Text = "" Then
'        MsgBox "Debe elegir la agencia", vbInformation, "Aviso"
'        cboAgencia.SetFocus
'        Exit Function
'    End If
    ValidaDatos = True
End Function

'COMENTADO POR APRI ERS036-2017
'Private Sub cmdExportar_Click()
'If Not (feDatos Is Nothing) Then
'    If feDatos.TextMatrix(1, 0) <> "" Then
'        Dim xlAplicacion As Excel.Application
'        Dim xlLibro As Excel.Workbook
'        Dim lbLibroOpen As Boolean
'        Dim lsArchivo As String
'        Dim lsHoja As String
'        Dim xlHoja1 As Excel.Worksheet
'        Dim xlHoja2 As Excel.Worksheet
'        Dim nLin As Long
'        Dim nItem As Long
'        Dim sColumna As String
'
'            lsArchivo = App.Path & "\SPOOLER\ReporteEnvioEstadoCuenta_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
'            lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
'            If Not lbLibroOpen Then
'                Exit Sub
'            End If
'            nLin = 1
'
'            If lnTipo = 1 Then
'                lsHoja = "Ahorros"
'                sColumna = "F"
'            Else
'                lsHoja = "Creditos"
'                sColumna = "R"
'            End If
'
'            pgb.value = 0
'            pgb.Min = 0
'
'            ExcelAddHoja lsHoja, xlLibro, xlHoja1
'
'            xlHoja1.Range("A1:Y1").EntireColumn.Font.FontStyle = "Arial"
'            xlHoja1.PageSetup.Orientation = xlLandscape
'            xlHoja1.PageSetup.CenterHorizontally = True
'            xlHoja1.PageSetup.Zoom = 75
'            xlHoja1.PageSetup.TopMargin = 2
'
'            xlHoja1.Range("A1:A1").RowHeight = 18
'            xlHoja1.Range("A1:A1").ColumnWidth = 5
'            xlHoja1.Range("B1:B1").ColumnWidth = 5
'            xlHoja1.Range("C1:C1").ColumnWidth = 40
'            xlHoja1.Range("D1:D1").ColumnWidth = 20
'            xlHoja1.Range("E1:E1").ColumnWidth = 40
'            If lnTipo = 1 Then
'                xlHoja1.Range("F1:F1").ColumnWidth = 5
'            Else
'                xlHoja1.Range("F1:Q1").ColumnWidth = 15
'                xlHoja1.Range("R1:R1").ColumnWidth = 5
'            End If
'            xlHoja1.Cells(1, 1) = " "
'
'            pgb.Max = feDatos.Rows - 1
'
'            xlHoja1.Cells(nLin, 1) = "REPORTE ENVIO ESTADO DE CUENTAS " & IIf(lnTipo = 1, "AHORROS", "CREDITOS")
'            xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).Merge True
'            xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).Font.Bold = True
'            xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).HorizontalAlignment = xlHAlignCenter
'        '    nLin = nLin + 1
'        '    xlHoja1.Cells(nLin, 2) = "Desde " & Right(psFecIni, 2) & "/" & Mid(psFecIni, 5, 2) & "/" & Left(psFecIni, 4) & _
'        '                             " Hasta " & Right(psFecFin, 2) & "/" & Mid(psFecFin, 5, 2) & "/" & Left(psFecFin, 4)
'        '    xlHoja1.Range("A" & nLin & ":Y" & nLin).Font.Bold = True
'        '    xlHoja1.Range("A" & nLin & ":Y" & nLin).Merge True
'        '    xlHoja1.Range("A" & nLin & ":Y" & nLin).HorizontalAlignment = xlHAlignCenter
'
'            nLin = nLin + 2
'
'            xlHoja1.Cells(nLin, 1) = "ITEM"
'            xlHoja1.Cells(nLin, 2) = "SEL"
'            xlHoja1.Cells(nLin, 3) = "CLIENTE"
'            xlHoja1.Cells(nLin, 4) = "CUENTA"
'            xlHoja1.Cells(nLin, 5) = "DIRECCION"
'            If lnTipo = 2 Then
'                xlHoja1.Cells(nLin, 6) = "CUOTAS PAG"
'                xlHoja1.Cells(nLin, 7) = "CUOTAS PEND"
'                xlHoja1.Cells(nLin, 8) = "ULTI CUOTA PAG"
'                xlHoja1.Cells(nLin, 9) = "CAPITAL"
'                xlHoja1.Cells(nLin, 10) = "INT COMP"
'                xlHoja1.Cells(nLin, 11) = "INT MORA"
'                xlHoja1.Cells(nLin, 12) = "GASTOS"
'                xlHoja1.Cells(nLin, 13) = "INT GRACIA"
'                xlHoja1.Cells(nLin, 14) = "INT SUSP"
'                xlHoja1.Cells(nLin, 15) = "FEC PAGO"
'                xlHoja1.Cells(nLin, 16) = "PROX CUOTA"
'                xlHoja1.Cells(nLin, 17) = "MONTO PAG"
'            End If
'            xlHoja1.Cells(nLin, IIf(lnTipo = 1, 6, 18)) = "PEND"
'
'            xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).Font.Bold = True
'            xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).HorizontalAlignment = xlHAlignCenter
'            xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
'            xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
'            xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).Borders(xlInsideVertical).Color = vbBlack
'            xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).Interior.Color = RGB(255, 50, 50)
'            xlHoja1.Range("A" & nLin & ":" & sColumna & nLin).Font.Color = RGB(255, 255, 255)
'
'            With xlHoja1.PageSetup
'                .LeftHeader = ""
'                .CenterHeader = ""
'                .RightHeader = ""
'                .LeftFooter = ""
'                .CenterFooter = ""
'                .RightFooter = ""
'
'                .PrintHeadings = False
'                .PrintGridlines = False
'                .PrintComments = xlPrintNoComments
'                .CenterHorizontally = True
'                .CenterVertically = False
'                .Orientation = xlLandscape
'                .Draft = False
'                .FirstPageNumber = xlAutomatic
'                .Order = xlDownThenOver
'                .BlackAndWhite = False
'                .Zoom = 55
'            End With
'
'            nItem = 1
'            nLin = nLin + 1
'            For nItem = 1 To feDatos.Rows - 1
'                xlHoja1.Range("A" & nLin & ":E" & nLin).HorizontalAlignment = xlHAlignLeft
'                xlHoja1.Range("I" & nLin & ":N" & nLin).NumberFormat = "#,##0.00"
'                xlHoja1.Range("Q" & nLin & ":Q" & nLin).NumberFormat = "#,##0.00"
'                xlHoja1.Cells(nLin, 1) = feDatos.TextMatrix(nItem, 0)
'                xlHoja1.Cells(nLin, 2) = IIf(feDatos.TextMatrix(nItem, 1) = "", "NO", "SI")
'                xlHoja1.Cells(nLin, 3) = feDatos.TextMatrix(nItem, 2)
'                xlHoja1.Cells(nLin, 4) = "'" & feDatos.TextMatrix(nItem, 3)
'                xlHoja1.Cells(nLin, 5) = feDatos.TextMatrix(nItem, 4)
'                If lnTipo = 2 Then
'                    xlHoja1.Cells(nLin, 6) = feDatos.TextMatrix(nItem, 5)
'                    xlHoja1.Cells(nLin, 7) = feDatos.TextMatrix(nItem, 6)
'                    xlHoja1.Cells(nLin, 8) = feDatos.TextMatrix(nItem, 7)
'                    xlHoja1.Cells(nLin, 9) = Format(feDatos.TextMatrix(nItem, 8), "#,##0.00")
'                    xlHoja1.Cells(nLin, 10) = Format(feDatos.TextMatrix(nItem, 9), "#,##0.00")
'                    xlHoja1.Cells(nLin, 11) = Format(feDatos.TextMatrix(nItem, 10), "#,##0.00")
'                    xlHoja1.Cells(nLin, 12) = Format(feDatos.TextMatrix(nItem, 11), "#,##0.00")
'                    xlHoja1.Cells(nLin, 13) = Format(feDatos.TextMatrix(nItem, 12), "#,##0.00")
'                    xlHoja1.Cells(nLin, 14) = Format(feDatos.TextMatrix(nItem, 13), "#,##0.00")
'                    xlHoja1.Cells(nLin, 15) = feDatos.TextMatrix(nItem, 14)
'                    xlHoja1.Cells(nLin, 16) = Format(feDatos.TextMatrix(nItem, 15), "#,##0.00")
'                    xlHoja1.Cells(nLin, 17) = feDatos.TextMatrix(nItem, 16)
'                End If
'                xlHoja1.Cells(nLin, IIf(lnTipo = 1, 6, 18)) = feDatos.TextMatrix(nItem, IIf(lnTipo = 1, 5, 17))
'                nLin = nLin + 1
'                pgb.value = pgb.value + 1
'            Next nItem
'
'            ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
'            CargaArchivo lsArchivo, App.Path & "\SPOOLER\"
'            pgb.value = 0
'            pgb.Min = 0
'    Else
'        MsgBox "No existen datos para exportar", vbInformation, "Aviso"
'    End If
'Else
'    MsgBox "No existen datos para exportar", vbInformation, "Aviso"
'End If
'End Sub

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
  MsgBox err.Description, vbInformation, "Aviso"
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
   MsgBox err.Description, vbInformation, "Aviso"
End Sub

'********************************
' Adiciona Hoja a LibroExcel
'********************************
Private Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet)
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = psHojName Then
       xlHoja1.Delete
       Exit For
    End If
Next
Set xlHoja1 = xlLibro.Worksheets.Add
xlHoja1.Name = psHojName
End Sub

Public Sub CargaArchivo(lsArchivo As String, lsRutaArchivo As String)
    Dim x As Long
    Dim Temp As String
    Temp = GetActiveWindow()
    x = ShellExecute(Temp, "open", lsArchivo, "", lsRutaArchivo, 1)
    If x <= 32 Then
        If x = 2 Then
            MsgBox "No se encuentra el Archivo adjunto, " & vbCr & " verifique el servidor de archivos", vbInformation, " Aviso "
        ElseIf x = 8 Then
            MsgBox "Memoria insuficiente ", vbInformation, " Aviso "
        Else
            MsgBox "No se pudo abrir el Archivo adjunto", vbInformation, " Aviso "
        End If
    End If
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
  KeyAscii = NumerosDecimales(txtAnio, KeyAscii, 15)
  If KeyAscii <> 13 Then Exit Sub
End Sub
