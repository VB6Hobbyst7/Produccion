Attribute VB_Name = "gColRecFunciones"
'**  Funciones Generales de Colocaciones-Recuperaciones.
'**
Option Explicit


Public Function fgIniciaAxCuentaRecuperaciones() As String
    fgIniciaAxCuentaRecuperaciones = gsCodCMAC & gsCodAge
End Function


Public Function fgEstadoColRecupDesc(ByVal pnEstado As Integer) As String
Dim lsDesc As String
    Select Case pnEstado
        Case gColocEstRecVigJud, gColocEstRecVigCast
            lsDesc = "Vigente"
        Case gColocEstRecCanJud, gColocEstRecCanCast
            lsDesc = "Cancelado"
    End Select
    fgEstadoColRecupDesc = lsDesc
End Function

Public Function fgCondicionColRecupDesc(ByVal pnEstado As Integer) As String
Dim lsDesc As String
    Select Case pnEstado
        Case gColocEstRecVigJud, gColocEstRecCanJud
            lsDesc = "Judicial"
        Case gColocEstRecVigCast, gColocEstRecCanCast
            lsDesc = "Castigado"
    End Select
    fgCondicionColRecupDesc = lsDesc
End Function

Public Function fgImprimeActuacionesProcesales(ByVal rs As ADODB.Recordset) As Boolean
    
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    

    Dim lbConexion As Boolean
    Dim lbExisteHoja  As Boolean
    Dim lsNomHoja As String
    Dim glsArchivo As String
    Dim liLineas As Integer
    Dim nNum As Double
    Dim fs As Scripting.FileSystemObject
    

    glsArchivo = "ActuacionesProcesales_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets.Add

    'xlHoja1.PageSetup.CenterHorizontally = True
    'xlHoja1.PageSetup.Zoom = 60
    'xlHoja1.PageSetup.Orientation = xlLandscape
    
    lbExisteHoja = False
    lsNomHoja = "ActuacionesProcesales"
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

    xlAplicacion.Range("A1:A1").ColumnWidth = 14
    xlAplicacion.Range("B1:B1").ColumnWidth = 30
    xlAplicacion.Range("C1:C1").ColumnWidth = 12
    xlAplicacion.Range("D1:D1").ColumnWidth = 12
    xlAplicacion.Range("E1:E1").ColumnWidth = 12
    xlAplicacion.Range("F1:F1").ColumnWidth = 12
    xlAplicacion.Range("G1:G1").ColumnWidth = 40
    xlAplicacion.Range("H1:H1").ColumnWidth = 15
    xlAplicacion.Range("I1:I1").ColumnWidth = 50
    xlAplicacion.Range("A1:Z2000").Font.Size = 8
    
    xlHoja1.Cells(1, 1) = gsNomCmac
    xlHoja1.Cells(2, 1) = gsNomAge
    xlHoja1.Cells(1, 8) = Trim(Format(gdFecSis, "dd/mm/yyyy hh:mm:ss"))
    xlHoja1.Cells(2, 8) = gsCodUser
    
    
    xlHoja1.Cells(4, 1) = "L I S T A D O  D E  A C T U A C I O N E S  P R O C E S A L E S"
    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(4, 9)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, 9)).Merge True
    xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, 9)).HorizontalAlignment = xlCenter
           
    liLineas = 6
    
    If Not (rs.EOF And rs.BOF) Then
          rs.MoveFirst
          xlHoja1.Cells(liLineas, 1) = "Credito"
          xlHoja1.Cells(liLineas, 2) = "Cliente"
          xlHoja1.Cells(liLineas, 3) = "Monto Prestamo"
          xlHoja1.Cells(liLineas, 4) = "Saldo Int. Comp."
          xlHoja1.Cells(liLineas, 5) = "Saldo Int. Mora."
          xlHoja1.Cells(liLineas, 6) = "Saldo de Gastos"
          xlHoja1.Cells(liLineas, 7) = "Comentario"
          xlHoja1.Cells(liLineas, 8) = "Fecha Vencimiento"
          xlHoja1.Cells(liLineas, 9) = "Abogado"
          xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 9)).Cells.Interior.Color = RGB(220, 220, 220)
          xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 9)).HorizontalAlignment = xlCenter
          liLineas = liLineas + 1
          Do Until rs.EOF
            xlHoja1.Cells(liLineas, 1) = rs!cCtaCod
            xlHoja1.Cells(liLineas, 2) = rs!Cliente
            xlHoja1.Cells(liLineas, 3) = ImpreFormat(rs!nMontoCol, 9, 2, True)
            xlHoja1.Cells(liLineas, 4) = ImpreFormat(rs!nSaldoIntComp, 9, 2, True)
            xlHoja1.Cells(liLineas, 5) = ImpreFormat(rs!nSaldoIntMor, 9, 2, True)
            xlHoja1.Cells(liLineas, 6) = ImpreFormat(rs!nSaldoGasto, 9, 2, True)
            nNum = CInt(Len(Replace(rs!cComenta, Chr(13), "")) / 100)
            If nNum = 0 Then
                nNum = 1
            End If
            nNum = nNum * 12.5
            xlHoja1.Cells(liLineas, 7) = Replace(rs!cComenta, Chr(13), "")
            xlHoja1.Range(xlHoja1.Cells(liLineas, 7), xlHoja1.Cells(liLineas, 7)).Merge True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 7), xlHoja1.Cells(liLineas, 7)).Cells.WrapText = True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 7), xlHoja1.Cells(liLineas, 7)).RowHeight = nNum
            
            xlHoja1.Cells(liLineas, 8) = Format(rs!dFechaAviso, "dd/mm/yyyy")
            xlHoja1.Cells(liLineas, 9) = rs!Abogado
            
            xlHoja1.Range(xlHoja1.Cells(liLineas - 1, 1), xlHoja1.Cells(liLineas, 9)).Borders.LineStyle = 1
            rs.MoveNext
            liLineas = liLineas + 1
           Loop
       
     
    '**************************************************************************************************
        
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        'Cierra el libro de trabajo
        xlLibro.Close
        'Cierra Microsoft Excel con el método Quit.
        xlAplicacion.Quit
        'Libera los objetos.
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        gFunContab.CargaArchivo glsArchivo, App.path & "\SPOOLER\"
        fgImprimeActuacionesProcesales = True
     Else
        fgImprimeActuacionesProcesales = False
     End If
End Function

'** DAOR 20070124
'** Visualizar actuaciones procesales en excel, de acuerdo al formato entregado por recuperaciones
Public Function fgImprimeActuacionesProcesales2(ByVal rs As ADODB.Recordset, psCondiciones As String) As Boolean
    
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    

    Dim lbConexion As Boolean
    Dim lbExisteHoja  As Boolean
    Dim lsNomHoja As String
    Dim glsArchivo As String
    Dim liLineas As Integer
    Dim nNum As Double
    Dim fs As Scripting.FileSystemObject
    Dim ClienteActual As String
    Dim sTempAbogado As String
    Dim sTempMoneda As String
    Dim nTotalMonedaCA As Double
    Dim nTotalMonedaSC As Double

    glsArchivo = "ActuacionesProcesales2_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets.Add

    'xlHoja1.PageSetup.CenterHorizontally = True
    'xlHoja1.PageSetup.Zoom = 60
    'xlHoja1.PageSetup.Orientation = xlLandscape
    
    lbExisteHoja = False
    lsNomHoja = "ActuacionesProcesales2"
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

    xlAplicacion.Range("A1:A1").ColumnWidth = 20
    xlAplicacion.Range("B1:B1").ColumnWidth = 20
    xlAplicacion.Range("C1:C1").ColumnWidth = 12
    xlAplicacion.Range("D1:D1").ColumnWidth = 12
    xlAplicacion.Range("E1:E1").ColumnWidth = 12
    xlAplicacion.Range("F1:F1").ColumnWidth = 8
    xlAplicacion.Range("G1:G1").ColumnWidth = 12
    xlAplicacion.Range("H1:H1").ColumnWidth = 10
    xlAplicacion.Range("I1:I1").ColumnWidth = 8
    xlAplicacion.Range("J1:J1").ColumnWidth = 30
    xlAplicacion.Range("K1:K1").ColumnWidth = 12
    xlAplicacion.Range("A1:Z2000").Font.Size = 8
    
    xlHoja1.Cells(1, 1) = gsNomCmac
    xlHoja1.Cells(2, 1) = gsNomAge
    xlHoja1.Cells(1, 9) = Trim(Format(gdFecSis, "dd/mm/yyyy hh:mm:ss"))
    xlHoja1.Cells(2, 9) = gsCodUser
    
    
    xlHoja1.Cells(4, 1) = "L I S T A D O  D E  A C T U A C I O N E S  P R O C E S A L E S"
    xlHoja1.Cells(5, 1) = psCondiciones
    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, 11)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, 11)).Merge True
    xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 11)).Merge True
    xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(5, 11)).HorizontalAlignment = xlCenter
           
    liLineas = 6
    
    If Not (rs.EOF And rs.BOF) Then
          rs.MoveFirst
          ClienteActual = ""
          sTempAbogado = "0"
          sTempMoneda = "0"
          nTotalMonedaCA = 0
          nTotalMonedaSC = 0
          Do Until rs.EOF
            If sTempAbogado <> rs!Abogado And sTempAbogado <> "0" Then sTempMoneda = ""
            If sTempMoneda <> rs!Moneda And sTempMoneda <> "0" Then
                xlHoja1.Cells(liLineas, 1) = " TOTAL "
                xlHoja1.Cells(liLineas, 4) = ImpreFormat(nTotalMonedaSC, 9, 2, True)
                xlHoja1.Cells(liLineas, 5) = ImpreFormat(nTotalMonedaCA, 9, 2, True)
                xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 1)).HorizontalAlignment = xlRight
                xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 11)).Font.Bold = True
                liLineas = liLineas + 1
            End If
            If sTempAbogado <> rs!Abogado Then
                sTempAbogado = rs!Abogado
                liLineas = liLineas + 1
                xlHoja1.Cells(liLineas, 1) = "Abogado : " & rs!Abogado
                xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 1)).Font.Bold = True
                liLineas = liLineas + 1
            End If
            If sTempMoneda <> rs!Moneda Then
                sTempMoneda = rs!Moneda
                nTotalMonedaCA = 0
                nTotalMonedaSC = 0
                xlHoja1.Cells(liLineas, 1) = "Moneda : " & IIf(rs!Moneda = "1", "Soles", "Dólares")
                xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 1)).Font.Bold = True
                liLineas = liLineas + 1
                xlHoja1.Cells(liLineas, 1) = "Demandado"
                xlHoja1.Cells(liLineas, 2) = "Juzgado"
                xlHoja1.Cells(liLineas, 3) = "Materia"
                xlHoja1.Cells(liLineas, 4) = "S. Capital"
                xlHoja1.Cells(liLineas, 5) = "C. Actual"
                xlHoja1.Cells(liLineas, 6) = "Fecha"
                xlHoja1.Cells(liLineas, 7) = "Exp. Nº"
                xlHoja1.Cells(liLineas, 8) = "Estado"
                xlHoja1.Cells(liLineas, 9) = "Fecha Act."
                xlHoja1.Cells(liLineas, 10) = "Actuacion Procesal"
                xlHoja1.Cells(liLineas, 11) = "Tipo"
                xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 11)).Cells.Interior.Color = RGB(220, 220, 220)
                xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 11)).HorizontalAlignment = xlCenter
                liLineas = liLineas + 1
            End If
            If (ClienteActual <> rs!Cliente) Then
                ClienteActual = rs!Cliente
                xlHoja1.Cells(liLineas, 1) = rs!Cliente & " - " & rs!cCtaCod
                xlHoja1.Cells(liLineas, 2) = rs!Juzgado
                xlHoja1.Cells(liLineas, 3) = rs!Materia
                xlHoja1.Cells(liLineas, 4) = ImpreFormat(rs!nSaldoTrans, 9, 2, True)
                xlHoja1.Cells(liLineas, 5) = ImpreFormat(rs!nSaldo, 9, 2, True)
                xlHoja1.Cells(liLineas, 6) = Format(rs!dIngRecup, "dd/mm/yyyy")
                xlHoja1.Cells(liLineas, 7) = rs!cNumExp
                xlHoja1.Cells(liLineas, 8) = rs!Estado
                nTotalMonedaCA = nTotalMonedaCA + rs!nSaldo
                nTotalMonedaSC = nTotalMonedaSC + rs!nSaldoTrans
            End If
            xlHoja1.Cells(liLineas, 9) = Format(rs!dFechaVencimiento, "dd/mm/yyyy")
            nNum = CInt(Len(Replace(rs!cComenta, Chr(13), "")) / 100)
            If nNum = 0 Then
                nNum = 1
            End If
            nNum = nNum * 12.5
            xlHoja1.Cells(liLineas, 10) = Replace(rs!cComenta, Chr(13), "")
            xlHoja1.Range(xlHoja1.Cells(liLineas, 10), xlHoja1.Cells(liLineas, 10)).Merge True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 10), xlHoja1.Cells(liLineas, 10)).Cells.WrapText = True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 10), xlHoja1.Cells(liLineas, 10)).RowHeight = nNum
            xlHoja1.Cells(liLineas, 11) = rs!TipoAct
            
            xlHoja1.Range(xlHoja1.Cells(liLineas - 1, 1), xlHoja1.Cells(liLineas, 11)).Borders.LineStyle = 1
            rs.MoveNext
            liLineas = liLineas + 1
           Loop
       
            xlHoja1.Cells(liLineas, 1) = " TOTAL "
            xlHoja1.Cells(liLineas, 4) = ImpreFormat(nTotalMonedaSC, 9, 2, True)
            xlHoja1.Cells(liLineas, 5) = ImpreFormat(nTotalMonedaCA, 9, 2, True)
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 1)).HorizontalAlignment = xlRight
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 11)).Font.Bold = True
     
    '**************************************************************************************************
        
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        'Cierra el libro de trabajo
        xlLibro.Close
        'Cierra Microsoft Excel con el método Quit.
        xlAplicacion.Quit
        'Libera los objetos.
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        gFunContab.CargaArchivo glsArchivo, App.path & "\SPOOLER\"
        fgImprimeActuacionesProcesales2 = True
     Else
        fgImprimeActuacionesProcesales2 = False
     End If
End Function






