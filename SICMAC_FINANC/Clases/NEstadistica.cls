VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NEstadistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub MuestraEstadisticaAdelantada(ByVal pdFecha As Date, ByVal pTipoCambio As Double)
    Dim m_Excel As Excel.Application
    Dim oLibroExcel As Excel.Workbook
    Dim odEstadisticas As DEstadistica
    Dim nTipoCambio As Double
    
    
    'Dim oHojaExcel As Excel.Worksheet
    Dim oCelda As Excel.Range
    
    Dim Progress As clsProgressBar
    Set Progress = New clsProgressBar
    Dim i As Long
    
    Set m_Excel = New Excel.Application
    m_Excel.Visible = True

    Set oLibroExcel = m_Excel.Workbooks.Add
       
    Select Case Month(pdFecha)
        Case 1
              Call ProcesarReporte(DateAdd("day", pdFecha, -1), pTipoCambio, oLibroExcel, "ENERO")
        Case 2
              Set odEstadisticas = New DEstadistica
              nTipoCambio = odEstadisticas.TipoCambio(1)
              Set odEstadisticas = Nothing
              Call ProcesarReporte("31/01/" & Year(pdFecha), pTipoCambio, oLibroExcel, "ENERO")
              Call ProcesarReporte(pdFecha, pTipoCambio, oLibroExcel, "FEBRERO", 1)
              'Call ProcesarReporte(DateAdd("day", pdFecha, -1), pTipoCambio, oLibroExcel, "FEBRERO", 1)
        Case 3
              Set odEstadisticas = New DEstadistica
              nTipoCambio = odEstadisticas.TipoCambio(1)
              Set odEstadisticas = Nothing
              Call ProcesarReporte("31/01/" & Year(pdFecha), nTipoCambio, oLibroExcel, "ENERO")
              
              Set odEstadisticas = New DEstadistica
              nTipoCambio = odEstadisticas.TipoCambio(2)
              Set odEstadisticas = Nothing
              Call ProcesarReporte("28/02/" & Year(pdFecha), nTipoCambio, oLibroExcel, "FEBRERO", 1)
              
              Call ProcesarReporte(pdFecha, pTipoCambio, oLibroExcel, "MARZO", 2)
              'Call ProcesarReporte(DateAdd("day", pdFecha, -1), pTipoCambio, oLibroExcel, "MARZO", 2)
              
        Case 4
              Set odEstadisticas = New DEstadistica
              Progress.Max = 100
              Progress.Progress 25, "Procesando Mes de Enero....."
              nTipoCambio = odEstadisticas.TipoCambio(1)
              Set odEstadisticas = Nothing
              Call ProcesarReporte("31/01/" & Year(pdFecha), nTipoCambio, oLibroExcel, "ENERO")
              
              Progress.Max = 100
              Progress.Progress 50, "Procesando Mes de Febrero....."
              Set odEstadisticas = New DEstadistica
              nTipoCambio = odEstadisticas.TipoCambio(2)
              Set odEstadisticas = Nothing
              Call ProcesarReporte("28/02/" & Year(pdFecha), nTipoCambio, oLibroExcel, "FEBRERO", 1)
              
              Progress.Max = 100
              Progress.Progress 75, "Procesando Mes de Marzo....."
              Set odEstadisticas = New DEstadistica
              nTipoCambio = odEstadisticas.TipoCambio(3)
              Set odEstadisticas = Nothing
              Call ProcesarReporte("31/03/" & Year(pdFecha), nTipoCambio, oLibroExcel, "MARZO", 2)
              
              Progress.Max = 100
              Progress.Progress 99, "Procesando Mes de Abril....."
              Call ProcesarReporte(pdFecha, pTipoCambio, oLibroExcel, "ABRIL", 3)
              'Call ProcesarReporte(DateAdd("d", pdFecha, -1), pTipoCambio, oLibroExcel, "ABRIL", 3)
              
        Case 5
              'Set odEstadisticas = New DEstadistica
              'Progress.Max = 100
              'Progress.Progress 25, "Procesando Mes de Enero....."
              'nTipoCambio = odEstadisticas.TipoCambio(1)
              'Set odEstadisticas = Nothing
              'Call ProcesarReporte("31/01/" & Year(pdFecha), nTipoCambio, oLibroExcel, "ENERO")
              
             ' Progress.Max = 100
             ' Progress.Progress 50, "Procesando Mes de Febrero....."
             ' Set odEstadisticas = New DEstadistica
             ' nTipoCambio = odEstadisticas.TipoCambio(2)
             ' Set odEstadisticas = Nothing
             ' Call ProcesarReporte("28/02/" & Year(pdFecha), nTipoCambio, oLibroExcel, "FEBRERO", 1)
              
             ' Progress.Max = 100
             ' Progress.Progress 56, "Procesando Mes de Marzo....."
             ' Set odEstadisticas = New DEstadistica
             ' nTipoCambio = odEstadisticas.TipoCambio(3)
             ' Set odEstadisticas = Nothing
             ' Call ProcesarReporte("31/03/" & Year(pdFecha), nTipoCambio, oLibroExcel, "MARZO", 2)
             '
'              Progress.Max = 100
'              Progress.Progress 75, "Procesando Mes de Abril....."
'              Set odEstadisticas = New DEstadistica
'              nTipoCambio = odEstadisticas.TipoCambio(4)
'              Set odEstadisticas = Nothing
'              Call ProcesarReporte("30/04/" & Year(pdFecha), nTipoCambio, oLibroExcel, "ABRIL", 3)
              
              Progress.Max = 100
              Progress.Progress 99, "Procesando Mes de Mayo....."
              Call ProcesarReporte(pdFecha, pTipoCambio, oLibroExcel, "MAYO", 4)
              'Call ProcesarReporte(DateAdd("d", pdFecha, -1), pTipoCambio, oLibroExcel, "MAYO", 4)
        
        Case 6
              'Set odEstadisticas = New DEstadistica
              'Progress.Max = 100
              'Progress.Progress 17, "Procesando Mes de Enero....."
              'nTipoCambio = odEstadisticas.TipoCambio(1)
              'Set odEstadisticas = Nothing
              'Call ProcesarReporte("31/01/" & Year(pdFecha), nTipoCambio, oLibroExcel, "ENERO")
              
'              Progress.Max = 100
'              Progress.Progress 34, "Procesando Mes de Febrero....."
'              Set odEstadisticas = New DEstadistica
'              nTipoCambio = odEstadisticas.TipoCambio(2)
'              Set odEstadisticas = Nothing
'              Call ProcesarReporte("28/02/" & Year(pdFecha), nTipoCambio, oLibroExcel, "FEBRERO", 1)
'
'              Progress.Max = 100
'              Progress.Progress 51, "Procesando Mes de Marzo....."
'              Set odEstadisticas = New DEstadistica
'              nTipoCambio = odEstadisticas.TipoCambio(3)
'              Set odEstadisticas = Nothing
'              Call ProcesarReporte("31/03/" & Year(pdFecha), nTipoCambio, oLibroExcel, "MARZO", 2)
'
'              Progress.Max = 100
'              Progress.Progress 68, "Procesando Mes de Abril....."
'              Set odEstadisticas = New DEstadistica
'              nTipoCambio = odEstadisticas.TipoCambio(4)
'              Set odEstadisticas = Nothing
'              Call ProcesarReporte("30/04/" & Year(pdFecha), nTipoCambio, oLibroExcel, "ABRIL", 3)
'
'              Progress.Max = 100
'              Progress.Progress 85, "Procesando Mes de Mayo....."
'              Set odEstadisticas = New DEstadistica
'              nTipoCambio = odEstadisticas.TipoCambio(5)
'              Set odEstadisticas = Nothing
'              Call ProcesarReporte("31/05/" & Year(pdFecha), nTipoCambio, oLibroExcel, "MAYO", 4)

              Progress.Max = 100
              Progress.Progress 99, "Procesando Mes de JUNIO....."
              Call ProcesarReporte(pdFecha, pTipoCambio, oLibroExcel, "JUNIO", 5)
              'Call ProcesarReporte(DateAdd("d", pdFecha, -1), pTipoCambio, oLibroExcel, "JUNIO", 5)
        
        Case 7
'              Set odEstadisticas = New DEstadistica
'              Progress.Max = 100
'              Progress.Progress 17, "Procesando Mes de Enero....."
'              nTipoCambio = odEstadisticas.TipoCambio(1)
'              Set odEstadisticas = Nothing
'              Call ProcesarReporte("31/01/" & Year(pdFecha), nTipoCambio, oLibroExcel, "ENERO")
'
'              Progress.Max = 100
'              Progress.Progress 34, "Procesando Mes de Febrero....."
'              Set odEstadisticas = New DEstadistica
'              nTipoCambio = odEstadisticas.TipoCambio(2)
'              Set odEstadisticas = Nothing
'              Call ProcesarReporte("28/02/" & Year(pdFecha), nTipoCambio, oLibroExcel, "FEBRERO", 1)
'
'              Progress.Max = 100
'              Progress.Progress 51, "Procesando Mes de Marzo....."
'              Set odEstadisticas = New DEstadistica
'              nTipoCambio = odEstadisticas.TipoCambio(3)
'              Set odEstadisticas = Nothing
'              Call ProcesarReporte("31/03/" & Year(pdFecha), nTipoCambio, oLibroExcel, "MARZO", 2)
'
'              Progress.Max = 100
'              Progress.Progress 68, "Procesando Mes de Abril....."
'              Set odEstadisticas = New DEstadistica
'              nTipoCambio = odEstadisticas.TipoCambio(4)
'              Set odEstadisticas = Nothing
'              Call ProcesarReporte("30/04/" & Year(pdFecha), nTipoCambio, oLibroExcel, "ABRIL", 3)
'
'              Progress.Max = 100
'              Progress.Progress 85, "Procesando Mes de Mayo....."
'              Set odEstadisticas = New DEstadistica
'              nTipoCambio = odEstadisticas.TipoCambio(5)
'              Set odEstadisticas = Nothing
'              Call ProcesarReporte("31/05/" & Year(pdFecha), nTipoCambio, oLibroExcel, "MAYO", 4)
              
              Progress.Max = 100
              Progress.Progress 85, "Procesando Mes de Junio....."
              Set odEstadisticas = New DEstadistica
              nTipoCambio = odEstadisticas.TipoCambio(6)
              Set odEstadisticas = Nothing
              Call ProcesarReporte("30/06/" & Year(pdFecha), nTipoCambio, oLibroExcel, "JUNIO", 5)
              
'              Progress.Max = 100
'              Progress.Progress 99, "Procesando Mes de Mayo....."
'              Call ProcesarReporte(DateAdd("d", pdFecha, -1), pTipoCambio, oLibroExcel, "JUNIO", 5)
        
        Case 8
              Progress.Max = 100
              Progress.Progress 85, "Procesando Mes de Julio....."
              Set odEstadisticas = New DEstadistica
              nTipoCambio = odEstadisticas.TipoCambio(7)
              Set odEstadisticas = Nothing
              Call ProcesarReporte("31/07/" & Year(pdFecha), nTipoCambio, oLibroExcel, "JULIO", 6)
        
        Case 9
              Progress.Max = 100
              Progress.Progress 85, "Procesando Mes de Agosto....."
              Set odEstadisticas = New DEstadistica
              nTipoCambio = odEstadisticas.TipoCambio(8)
              Set odEstadisticas = Nothing
              Call ProcesarReporte("31/08/" & Year(pdFecha), nTipoCambio, oLibroExcel, "AGOSTO", 7)
        
        Case 10
              Progress.Max = 100
              Progress.Progress 85, "Procesando Mes de Septiembre....."
              Set odEstadisticas = New DEstadistica
              nTipoCambio = odEstadisticas.TipoCambio(9)
              Set odEstadisticas = Nothing
              Call ProcesarReporte("30/09/" & Year(pdFecha), nTipoCambio, oLibroExcel, "SEPTIEMBRE", 8)
        
        Case 11
              Progress.Max = 100
              Progress.Progress 85, "Procesando Mes de Octubre....."
              Set odEstadisticas = New DEstadistica
              nTipoCambio = odEstadisticas.TipoCambio(10)
              Set odEstadisticas = Nothing
              Call ProcesarReporte("31/10/" & Year(pdFecha), nTipoCambio, oLibroExcel, "OCTUBRE", 9)
        
        Case 12
              Progress.Max = 100
              Progress.Progress 85, "Procesando Mes de Noviembre....."
              Set odEstadisticas = New DEstadistica
              nTipoCambio = odEstadisticas.TipoCambio(11)
              Set odEstadisticas = Nothing
              Call ProcesarReporte("30/11/" & Year(pdFecha), nTipoCambio, oLibroExcel, "NOVIEMBRE", 10)
    
        Case 13
              Progress.Max = 100
              Progress.Progress 85, "Procesando Mes de Diciembre....."
              Set odEstadisticas = New DEstadistica
              nTipoCambio = odEstadisticas.TipoCambio(12)
              Set odEstadisticas = Nothing
              Call ProcesarReporte("31/12/" & Year(pdFecha), nTipoCambio, oLibroExcel, "DICIEMBRE", 11)
    
    End Select
    
    oLibroExcel.PrintPreview
End Sub

Sub ProcesarReporte(ByVal pdFecha As String, ByVal pnTipoCambio, _
ByRef pLibroExcel As Excel.Workbook, ByVal pNombreMes As String, Optional ByVal pNumHoja As Integer = 0)

    Dim oHojaExcel As Excel.Worksheet
    Dim nFila As Integer
    Dim odEstadisticas As DEstadistica
    Dim rs As ADODB.Recordset
    Dim nSumaSoles As Double
    Dim nSumaDolares As Double
    
'    If pNumHoja <> 0 Then
'        pLibroExcel.Worksheets(pNumHoja - 1).Select
'    End If
    
   ' If pNumHoja = 0 Then
        Set oHojaExcel = pLibroExcel.Worksheets.Add
  '  Else
   '     Set oHojaExcel = pLibroExcel.Worksheets(1).Add
  '  End If
    
    oHojaExcel.Name = pNombreMes
    oHojaExcel.PageSetup.Zoom = 80
    oHojaExcel.PageSetup.Orientation = xlLandscape
    
    'oHojaExcel.Columns("A:A").EntireColumn.AutoFit
    
    oHojaExcel.Columns("B:B").NumberFormat = "0.00"
    oHojaExcel.Columns("C:C").NumberFormat = "0.00"
    oHojaExcel.Columns("D:D").NumberFormat = "0.00"
    oHojaExcel.Columns("E:E").NumberFormat = "0.00"
    
    oHojaExcel.Range("A4:E4").Merge
    oHojaExcel.Range("A4:E4").value = "ESTADISTICA ADELANTADA AL " & pdFecha
    oHojaExcel.Range("A4:E4").Font.Italic = True
    oHojaExcel.Range("A4:E4").Font.Size = 13
    oHojaExcel.Range("A4:E4").HorizontalAlignment = xlCenter
    oHojaExcel.Range("A4:E4").Font.Bold = True
    
    oHojaExcel.Range("A5:E5").Merge
    oHojaExcel.Range("A5:E5").value = "(En Nuevos Soles)"
    oHojaExcel.Range("A5:E5").Font.Italic = True
    oHojaExcel.Range("A5:E5").Font.Size = 12
    oHojaExcel.Range("A5:E5").HorizontalAlignment = xlCenter
    oHojaExcel.Range("A5:E5").Font.Bold = True
    
    oHojaExcel.Range("A6:E6").Merge
    oHojaExcel.Range("A6:E6").value = "Institucion : EDPYME CONFIANZA S.A."
    oHojaExcel.Range("A6:E6").Font.Italic = True
    oHojaExcel.Range("A6:E6").Font.Size = 11
    oHojaExcel.Range("A6:E6").HorizontalAlignment = xlCenter
    oHojaExcel.Range("A6:E6").Font.Bold = True
    
    'oHojaExcel.Cells(3, 4) = "(en nuevo soles)"
    'oHojaExcel.Cells(3, 4).Font.Bold = True
    
    'oHojaExcel.Cells(4, 4) = "Institucion Caja Municipal de Ahorro y Credito de ICA.SA"
    'oHojaExcel.Cells(4, 4).Font.Bold = True
    
    'Configurando los bordes
    
    With oHojaExcel.Range("A11:E48").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With oHojaExcel.Range("A11:E48").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With oHojaExcel.Range("A11:E48").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With oHojaExcel.Range("A11:E48").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With oHojaExcel.Range("A11:E48").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With oHojaExcel.Range("A11:E48").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    'Armando Columnas 'Caja General'
     nFila = 11
     With oHojaExcel
        .Cells(nFila, 1) = "CUENTAS"
        .Cells(nFila, 1).Font.Bold = True
        .Cells(nFila, 2) = "Monto M/N"
        .Cells(nFila, 2).Font.Bold = True
        .Cells(nFila, 3) = " Tasa M/N "
        .Cells(nFila, 3).Font.Bold = True
        .Cells(nFila, 4) = "Monto M/E"
        .Cells(nFila, 4).Font.Bold = True
        .Cells(nFila, 5) = " Tasa M/E "
        .Cells(nFila, 5).Font.Bold = True
    
     'nFila = nFila + 1 '12
     End With
     
     nFila = 12
     'Armando Filas de Caja
     With oHojaExcel
        .Cells(12, 1) = "Caja"
        .Cells(12, 1).Font.Bold = True
        
        .Cells(13, 1) = "Dep en Cta Cte (*)"
        .Cells(13, 1).Font.Bold = True
        
        .Cells(14, 1) = "Deposito de Ahorro"
        .Cells(14, 1).Font.Bold = True
        
        .Cells(15, 1) = "Dep a Plazo"
        .Cells(15, 1).Font.Bold = True
        
        .Cells(16, 1) = "Dep en CMAC (Aho y Plazo)"
        .Cells(16, 1).Font.Bold = True
        
        .Cells(17, 1) = "TOTAl DISPONIBLE"
        .Cells(17, 1).Font.Bold = True
        
        Set odEstadisticas = New DEstadistica
        Set rs = odEstadisticas.Recup_Caja(CDate(pdFecha), pnTipoCambio)
                
        Set odEstadisticas = Nothing
        
        Do Until rs.EOF
            If Trim(rs!Cta) = "Caja" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(12, 2) = rs!Caja
                Else
                    .Cells(12, 4) = rs!Caja
                End If
            ElseIf Trim(rs!Cta) = "Dep. en Cta Cte." Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(13, 2) = rs!Caja
                Else
                    .Cells(13, 4) = rs!Caja
                End If
            ElseIf Trim(rs!Cta) = "Depósitos Ahorro" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(14, 2) = rs!Caja
                Else
                    .Cells(14, 4) = rs!Caja
                End If
            ElseIf Trim(rs!Cta) = "Dep. a plazo" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(15, 2) = rs!Caja
                Else
                    .Cells(15, 4) = rs!Caja
                End If
            ElseIf Trim(rs!Cta) = "Dep. en CMAC (Aho y Plazo)" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(16, 2) = rs!Caja
                Else
                    .Cells(16, 4) = rs!Caja
                End If
            End If
            If Trim(rs!cMoneda) = "Soles" Then
                nSumaSoles = nSumaSoles + rs!Caja
            Else
                nSumaDolares = nSumaDolares + rs!Caja
            End If
            rs.MoveNext
        Loop
        Set rs = Nothing
         .Cells(17, 2) = nSumaSoles
         .Cells(17, 4) = nSumaDolares
     End With
     
    'Armando Columnas Inversiones Financieras
     nFila = 18
     With oHojaExcel
     
        .Cells(18, 1) = "INVERSIONES FINANCIERAS"
        .Cells(18, 1).Font.Bold = True
     
        .Cells(19, 1) = "Credito Prendario"
        .Cells(19, 1).Font.Bold = True
         
        .Cells(20, 1) = "Credito Pequeña Empresa"
        .Cells(20, 1).Font.Bold = True
        
        .Cells(21, 1) = "Credito Personal"
        .Cells(21, 1).Font.Bold = True
        
        .Cells(22, 1) = "Credito Agricola"
        .Cells(22, 1).Font.Bold = True
        
        .Cells(23, 1) = "Credito CTS"
        .Cells(23, 1).Font.Bold = True
        
        .Cells(24, 1) = "Prestamo CMAC"
        .Cells(24, 1).Font.Bold = True
         
        .Cells(25, 1) = "TOTAL COLOCACIONES"
        .Cells(25, 1).Font.Bold = True
      
      'INVERSIONES FINANCIERAS
        Set odEstadisticas = New DEstadistica
        Set rs = odEstadisticas.Recup_Inversiones(pdFecha, pnTipoCambio)
        Set odEstadisticas = Nothing
        Do Until rs.EOF
            If Trim(rs!Cuentas) = "Inversiones" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(18, 2) = rs!nValor
                Else
                    .Cells(18, 4) = rs!nValor
                End If
            End If
            rs.MoveNext
        Loop
        Set rs = Nothing
         
      'Llenando las informacion de creditos
        Set odEstadisticas = New DEstadistica
        Set rs = odEstadisticas.RecupCreditoByFecha(Format(pdFecha, "MM/dd/yyyy"), pnTipoCambio)
        Set odEstadisticas = Nothing
        nSumaSoles = 0
        nSumaDolares = 0
        Do Until rs.EOF
            
            If rs!cDescripcion = "Credito Agricola" Then
                If rs!cMoneda = "Soles" Then
                   .Cells(22, 2) = rs!nMonto
                   .Cells(22, 3) = rs!TasaPromedio
                Else
                   .Cells(22, 4) = Format(rs!nMonto, "#0.00")
                   .Cells(22, 5) = rs!TasaPromedio
                End If
            ElseIf rs!cDescripcion = "Prendario " Then
                If rs!cMoneda = "Soles" Then
                   .Cells(19, 2) = rs!nMonto
                   .Cells(19, 3) = rs!TasaPromedio
                Else
                   .Cells(19, 4) = rs!nMonto
                   .Cells(19, 5) = rs!TasaPromedio
                End If
            ElseIf rs!cDescripcion = "Credito Pequeña Empresa" Then
                If rs!cMoneda = "Soles" Then
                   .Cells(20, 2) = rs!nMonto
                   .Cells(20, 3) = rs!TasaPromedio
                Else
                   .Cells(20, 4) = rs!nMonto
                   .Cells(20, 5) = rs!TasaPromedio
                End If
            ElseIf rs!cDescripcion = "Credito Personal" Then
                 If rs!cMoneda = "Soles" Then
                   .Cells(21, 2) = rs!nMonto
                   .Cells(21, 3) = rs!TasaPromedio
                Else
                   .Cells(21, 4) = rs!nMonto
                   .Cells(21, 5) = rs!TasaPromedio
                End If
            ElseIf rs!cDescripcion = "Credito CTS" Then
               If rs!cMoneda = "Soles" Then
                   .Cells(23, 2) = rs!nMonto
                   .Cells(23, 3) = rs!TasaPromedio
                Else
                   .Cells(23, 4) = rs!nMonto
                   .Cells(23, 5) = rs!TasaPromedio
                End If
            
           ElseIf rs!cDescripcion = "Prestamos CMAC" Then
                If rs!cMoneda = "Soles" Then
                   .Cells(24, 2) = rs!nMonto
                   .Cells(24, 3) = rs!TasaPromedio
                Else
                   .Cells(24, 4) = rs!nMonto
                   .Cells(24, 5) = rs!TasaPromedio
                End If
           End If
            If Trim(rs!cMoneda) = "Soles" Then
                nSumaSoles = nSumaSoles + rs!nMonto
            Else
                nSumaDolares = nSumaDolares + rs!nMonto
            End If
           rs.MoveNext
        Loop
        Set rs = Nothing
        .Cells(25, 2) = nSumaSoles
        .Cells(25, 4) = nSumaDolares
     End With
         
     nFila = 26
     
     'Depositos
     With oHojaExcel
         .Cells(26, 1) = "Ahorro Corriente" '21
         .Cells(26, 1).Font.Bold = True
         
         .Cells(27, 1) = "Plazo Fijo" '22
         .Cells(27, 1).Font.Bold = True
         
         .Cells(28, 1) = "Dep.Aho.FONCODES" '23
         .Cells(28, 1).Font.Bold = True
         
         .Cells(29, 1) = "Dep.FOCMAN P.F" '24
         .Cells(29, 1).Font.Bold = True
         
         .Cells(30, 1) = "Dep de CMAC Ahorro" '25
         .Cells(30, 1).Font.Bold = True
         
         .Cells(31, 1) = "Dep de CMAC Plazo Fijo" '28
         .Cells(31, 1).Font.Bold = True
         
         .Cells(32, 1) = "Dep de CTS" '28
         .Cells(32, 1).Font.Bold = True
         
         .Cells(33, 1) = "TOTAL DEPOSITOS" '29
         .Cells(33, 1).Font.Bold = True

        Set odEstadisticas = New DEstadistica
        Set rs = odEstadisticas.GeneraRepCapEAFecha(pdFecha, pnTipoCambio)
        Set odEstadisticas = Nothing
        nSumaSoles = 0
        nSumaDolares = 0
        Do Until rs.EOF
            If Trim(rs!Reporte) = "AHORRO CORRIENTE" Then
                    .Cells(26, 2) = rs!TotalSoles
                    .Cells(26, 3) = rs!TasaPromSol
                    .Cells(26, 4) = rs!TotalSolesME
                    .Cells(26, 5) = rs!TasaPromDol
            ElseIf Trim(rs!Reporte) = "PLAZO FIJO" Then
                    .Cells(27, 2) = rs!TotalSoles
                    .Cells(27, 3) = rs!TasaPromSol
                    .Cells(27, 4) = rs!TotalSolesME
                    .Cells(27, 5) = rs!TasaPromDol
            ElseIf Trim(rs!Reporte) = "DEP. AHO. FONCODES" Then
                    .Cells(28, 2) = rs!TotalSoles
                    .Cells(28, 3) = rs!TasaPromSol
                    .Cells(28, 4) = rs!TotalSolesME
                    .Cells(28, 5) = rs!TasaPromDol
            ElseIf Trim(rs!Reporte) = "DEP. FOCMAN PF." Then
                    .Cells(29, 2) = rs!TotalSoles
                    .Cells(29, 3) = rs!TasaPromSol
                    .Cells(29, 4) = rs!TotalSolesME
                    .Cells(29, 5) = rs!TasaPromDol
            ElseIf Trim(rs!Reporte) = "DEP. CMAC AHORRO" Then
                    .Cells(30, 2) = rs!TotalSoles
                    .Cells(30, 3) = rs!TasaPromSol
                    .Cells(30, 4) = rs!TotalSolesME
                    .Cells(30, 5) = rs!TasaPromDol
            ElseIf Trim(rs!Reporte) = "DEP. CMAC PLAZOFIJO" Then
                    .Cells(31, 2) = rs!TotalSoles
                    .Cells(31, 3) = rs!TasaPromSol
                    .Cells(31, 4) = rs!TotalSolesME
                    .Cells(31, 5) = rs!TasaPromDol
            ElseIf Trim(rs!Reporte) = "DEP. CTS" Then
                    .Cells(32, 2) = rs!TotalSoles
                    .Cells(32, 3) = rs!TasaPromSol
                    .Cells(32, 4) = rs!TotalSolesME
                    .Cells(32, 5) = rs!TasaPromDol
            End If
            nSumaSoles = nSumaSoles + rs!TotalSoles
            nSumaSoles = nSumaSoles + rs!TotalSolesME
            rs.MoveNext
        Loop
        .Cells(33, 2) = nSumaSoles
        .Cells(33, 4) = nSumaDolares
    
    End With
    Set rs = Nothing
'----

     nFila = 34
     
     'Adeudados
     
     nSumaSoles = 0
     nSumaDolares = 0
     With oHojaExcel
        .Cells(34, 1) = "Prest.Banco Wiesse Sudameris"
        .Cells(34, 1).Font.Bold = True
        
        .Cells(35, 1) = "Prest.Bancos CREDITO"
        .Cells(35, 1).Font.Bold = True
        
        .Cells(36, 1) = "Prest.Bancos AGROBANCO"
        .Cells(36, 1).Font.Bold = True
        
        .Cells(37, 1) = "Prestamos de COFIDE"
        .Cells(37, 1).Font.Bold = True
        
        .Cells(38, 1) = "Prestamos de FOCMAC"
        .Cells(38, 1).Font.Bold = True
        
        .Cells(39, 1) = "Prestamos de BID"
        .Cells(39, 1).Font.Bold = True
        
        .Cells(40, 1) = "Prestamos de FONCODES(En Admin)"
        .Cells(40, 1).Font.Bold = True
        
        .Cells(41, 1) = "Prestamos de otra CMAC"
        .Cells(41, 1).Font.Bold = True
        
        .Cells(42, 1) = "Fondos en Fideicomiso/Adm."
        .Cells(42, 1).Font.Bold = True

        .Cells(43, 1) = "TOTAL ADEUDADOS"
        .Cells(43, 1).Font.Bold = True
        
        '.Cells(44, 1) = "PATRIMONIO"
       ' .Cells(44, 1).Font.Bold = True
        
        Set odEstadisticas = New DEstadistica
        Set rs = odEstadisticas.Recup_Adeudados(pdFecha, pnTipoCambio)
        Set odEstadisticas = Nothing
        
        Do Until rs.EOF
            If Trim(rs!Cuentas) = "Prest. Bancos Wiesse Sud" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(34, 2) = rs!nValor
                Else
                    .Cells(34, 4) = rs!nValor
                End If
            ElseIf Trim(rs!Cuentas) = "Prest. Bancos Credito" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(35, 2) = rs!nValor
                Else
                    .Cells(35, 4) = rs!nValor
                End If
            ElseIf Trim(rs!Cuentas) = "Prest. Bancos AgroBanco" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(36, 2) = rs!nValor
                Else
                    .Cells(36, 4) = rs!nValor
                End If
            ElseIf Trim(rs!Cuentas) = "Prest. Bancos COFIDE" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(37, 2) = rs!nValor
                Else
                    .Cells(37, 4) = rs!nValor
                End If
            ElseIf Trim(rs!Cuentas) = "Prestamos de FOCMAC" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(38, 2) = rs!nValor
                Else
                    .Cells(38, 4) = rs!nValor
                End If
            ElseIf Trim(rs!Cuentas) = "Prestamos de BID" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(39, 2) = rs!nValor
                Else
                    .Cells(39, 4) = rs!nValor
                End If
            
            ElseIf Trim(rs!Cuentas) = "Prestamos de FONCODES" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(40, 2) = rs!nValor
                Else
                    .Cells(40, 4) = rs!nValor
                End If
            ElseIf Trim(rs!Cuentas) = "Prestamos de Otra CMAC" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(41, 2) = rs!nValor
                Else
                    .Cells(41, 4) = rs!nValor
                End If
            ElseIf Trim(rs!Cuentas) = "Fondos en Fideicomiso" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(42, 2) = rs!nValor
                Else
                    .Cells(42, 4) = rs!nValor
                End If
            End If
        
            If Trim(rs!cMoneda) = "Soles" Then
                nSumaSoles = nSumaSoles + rs!nValor
            Else
                nSumaDolares = nSumaDolares + rs!nValor
            End If
            rs.MoveNext
        Loop
        .Cells(43, 2) = nSumaSoles
        .Cells(43, 4) = nSumaDolares
     End With
     
     nSumaSoles = 0
     nSumaDolares = 0
     'Gastos
     nFila = 45
     With oHojaExcel
        .Cells(45, 1) = "PATRIMONIO"
        .Cells(45, 1).Font.Bold = True
     
        .Cells(46, 1) = "Gastos  de Personal"
        .Cells(46, 1).Font.Bold = True
        
        .Cells(47, 1) = "Gastos Administrativos"
        .Cells(47, 1).Font.Bold = True
        
        Set odEstadisticas = New DEstadistica
        Set rs = odEstadisticas.Recup_Patrimonio(pdFecha, pnTipoCambio)
        Set odEstadisticas = Nothing
        
        Do Until rs.EOF
        
        If Trim(rs!Cuentas) = "Patrimonio" Then
            If Trim(rs!cMoneda) = "Soles" Then
                .Cells(45, 2) = rs!Patrimonio
            Else
                .Cells(45, 4) = rs!Patrimonio
            End If
        End If
         rs.MoveNext
        Loop
        Set rs = Nothing
        
        Set odEstadisticas = New DEstadistica
        Set rs = odEstadisticas.Recup_Gastos(pdFecha, pnTipoCambio)
        Set odEstadisticas = Nothing
        
        Do Until rs.EOF
            If Trim(rs!Cuentas) = "Gastos Administrativos" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(47, 2) = rs!nValor
                Else
                    .Cells(47, 4) = rs!nValor
                End If
                
            ElseIf Trim(rs!Cuentas) = "Gastos de Personal" Then
                If Trim(rs!cMoneda) = "Soles" Then
                    .Cells(46, 2) = rs!nValor
                Else
                    .Cells(46, 4) = rs!nValor
                End If
            End If
            rs.MoveNext
        Loop
        Set rs = Nothing
     End With
     
     
    With oHojaExcel.Range("A49:E70").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With oHojaExcel.Range("A49:E70").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With oHojaExcel.Range("A49:E70").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With oHojaExcel.Range("A49:E70").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With oHojaExcel.Range("A49:E70").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With oHojaExcel.Range("A49:E70").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
     
     
     
     nFila = 49
     
     With oHojaExcel
        .Cells(49, 1) = "OPERACIONES ESTIMADAS AL 31 DE ENERO   DEL 2005"
        .Cells(49, 1).Font.Bold = True
     
        .Cells(50, 1) = "PEQUEÑA EMPRESA"
        .Cells(50, 1).Font.Bold = True
        
        .Cells(51, 1) = "Nº de prést.otorg.acum. Peq. empresa "
        .Cells(51, 1).Font.Bold = True
        
        .Cells(52, 1) = "Nº préstamos vigentes Peq. Empresa"
        .Cells(52, 1).Font.Bold = True
         
        .Cells(53, 1) = "Monto Otorg. Acum. Peq. Empresa"
        .Cells(53, 1).Font.Bold = True

        .Cells(54, 1) = "AGRICOLA"
        .Cells(54, 1).Font.Bold = True
         
        .Cells(55, 1) = "Nº de prést.otorg.acum. Agricola"
        .Cells(55, 1).Font.Bold = True
        
        .Cells(56, 1) = "Nº préstamos vigentes Agricola"
        .Cells(56, 1).Font.Bold = True
         
        .Cells(57, 1) = "Monto Otorg. Acum. Agricola"
        .Cells(57, 1).Font.Bold = True
        
        .Cells(58, 1) = "PRENDARIO"
        .Cells(58, 1).Font.Bold = True
         
        .Cells(59, 1) = "Nº de prést.Otorg. acum.Prendario"
        .Cells(59, 1).Font.Bold = True
         
        .Cells(60, 1) = "Nº préstamos vigentes Prendario"
        .Cells(60, 1).Font.Bold = True
         
        .Cells(61, 1) = "Monto Otorgado acumulado Prendario"
        .Cells(61, 1).Font.Bold = True
         
        .Cells(62, 1) = "PERSONALES"
        .Cells(62, 1).Font.Bold = True
         
        .Cells(63, 1) = "Nºprést.otorg.acum.Dcto planilla,CTS Perso."
        .Cells(63, 1).Font.Bold = True
        
        .Cells(64, 1) = "Nº préstamo Dcto planilla vigente CTS Pers"
        .Cells(64, 1).Font.Bold = True
         
        .Cells(65, 1) = "Monto otorg.acum.Créd.dcto planilla CTS Per"
        .Cells(65, 1).Font.Bold = True
        
        .Cells(66, 1) = "Valor tasación oro 18K"
        .Cells(66, 1).Font.Bold = True
                 
        .Cells(67, 1) = "Porcentaje de Tasación para préstamo"
        .Cells(67, 1).Font.Bold = True
                 
        .Cells(68, 1) = "Depósitos del público"
        .Cells(68, 1).Font.Bold = True
        
        .Cells(69, 1) = "Depósitos total Concejo Provincial"
        .Cells(69, 1).Font.Bold = True
        
        .Cells(70, 1) = "Depósitos con Distritod. Emp.Mun.Reg. Focodes .."
        .Cells(70, 1).Font.Bold = True
        
        .Cells(71, 1) = "Depósitos otras Cmacs."
        .Cells(71, 1).Font.Bold = True
        
        .Cells(74, 1) = "TIPO DE CAMBIO"
        .Cells(74, 1).Font.Bold = True
        
        .Cells(75, 1) = "TOTAL PERSONAL"
        .Cells(75, 1).Font.Bold = True
        
        Set odEstadisticas = New DEstadistica
        Set rs = odEstadisticas.RecupConsolidado(Format(pdFecha, "MM/dd/yyyy"), pnTipoCambio)
        Set odEstadisticas = Nothing
        
        Do Until rs.EOF
            If rs!cProducto = "Pymes" Then
                If Trim(rs!cDescripcion) = "Desembolsos del Año" Then
                    .Cells(51, 2) = rs!nValor
                ElseIf Trim(rs!cDescripcion) = "Prestamos Vigentes" Then
                    .Cells(52, 2) = rs!nValor
                ElseIf Trim(rs!cDescripcion) = "Monto Acumulado" Then
                    .Cells(53, 2) = rs!nValor
                End If
             ElseIf rs!cProducto = "Agricolas" Then
                If Trim(rs!cDescripcion) = "Desembolsos del Año" Then
                    .Cells(55, 2) = rs!nValor
                ElseIf Trim(rs!cDescripcion) = "Prestamo Vigentes" Then
                    .Cells(56, 2) = rs!nValor
                ElseIf Trim(rs!cDescripcion) = "Monto Acumulado" Then
                    .Cells(57, 2) = rs!nValor
                End If
             ElseIf Trim(rs!cProducto) = "Prendario" Then
                If Trim(rs!cDescripcion) = "Desembolsos del Año" Then
                    .Cells(59, 2) = rs!nValor
                ElseIf Trim(rs!cDescripcion) = "Prestamos Vigentes" Then
                    .Cells(60, 2) = rs!nValor
                ElseIf Trim(rs!cDescripcion) = "Monto Acumulado" Then
                    .Cells(61, 2) = rs!nValor
                End If
             ElseIf Trim(rs!cProducto) = "Consumo" Then
                If Trim(rs!cDescripcion) = "Desembolsos del Año" Then
                    .Cells(63, 2) = rs!nValor
                ElseIf Trim(rs!cDescripcion) = "Prestamos Vigentes" Then
                    .Cells(64, 2) = rs!nValor
                ElseIf Trim(rs!cDescripcion) = "Monto Acumulado" Then
                    .Cells(66, 2) = rs!nValor
                End If
             End If
             
            rs.MoveNext
        Loop
        
        Set odEstadisticas = New DEstadistica
        Set rs = odEstadisticas.GeneraRepCapEAFechaConsol(pdFecha, pnTipoCambio)
        Set odEstadisticas = Nothing
        
       Do Until rs.EOF
            If rs!Reporte = "DEPOSITOS DEL PUBLICO" Then
                    .Cells(68, 2) = rs!Monto
             ElseIf rs!Reporte = "DEPOSITOS TOTAL CONCEJO PROVINCIAL" Then
                    .Cells(69, 2) = rs!Monto
             ElseIf rs!Reporte = "DEPOSITOS DISTRITO EMP.MUN.REG" Then
                    .Cells(70, 2) = rs!Monto
             ElseIf rs!Reporte = "DEPOSITOS DE CMACS" Then
                    .Cells(71, 2) = rs!Monto
             End If
             
            rs.MoveNext
        Loop
        Set rs = Nothing
       
        Set odEstadisticas = New DEstadistica
        Set rs = odEstadisticas.Recup_Personal(pdFecha, pnTipoCambio)
        Set odEstadisticas = Nothing
            
            .Cells(74, 2) = nTipoCambio
            .Cells(75, 2) = rs!Personal
        
        oHojaExcel.Columns("A:A").EntireColumn.AutoFit
        Set rs = Nothing
        
     End With
     
End Sub



