Attribute VB_Name = "gFunExcelCAP"
Public Sub ImpExportarExcelCapta(pnMonto As Long, pnTipoCambio As Double, psCodAge As String, psNomCmac As String, psNomAge As String, psFecSis As Date, psResp As String, pdFecha As Date, pdFechaF As Date, psOpe As String)
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim nFila As Long, i As Long

Dim orep As cOMNCaptaGenerales.NCOMCaptaReportes
Dim rsTemp As ADODB.Recordset


If MsgBox("Este reporte puede demorar unos minutos..." & vbCrLf & "¿Desea procesar la información ?", vbYesNo + vbQuestion, "AVISO") = vbNo Then
    Exit Sub
End If


Set orep = New cOMNCaptaGenerales.NCOMCaptaReportes
'(lsRep, Me.TxtFecha, Me.txtFechaF, Me.txtMonto.value, Me.txtMontoF.value, gsNomAge, gsNomCmac, gdFecSis, Me.TxtAgencia.Text, TxtBuscarUser.Text, rtfCartas.Text, Val(EditMoney3.Text), lsEstadosCheques, lsOptionsCheques, lsOrden, lscheck, lscmacllamada, lscmacrecepcion, pspersoneria)
'  lscadena = lscadena & GetRepCapNMejoresClientes(CLng(pnMontoIni), pnTipoCambio, psCodAge, psEmpresa, psNomAge, pdFecSis)
Select Case psOpe
       Case "280708"
            'Set rsTemp = orep.GetDataCapNMejoresClientes(Me.txtMonto.value, Val(EditMoney3.Text), Me.TxtAgencia.Text, gsNomCmac, gsNomAge, gdFecSis, 1)
            Set rsTemp = orep.GetDataCapNMejoresClientes(pnMonto, pnTipoCambio, psCodAge, psNomCmac, psNomAge, psFecSis, 1)
       Case "280709"
            Set rsTemp = orep.GetDataCapNMejoresClientes(pnMonto, pnTipoCambio, psCodAge, psNomCmac, psNomAge, psFecSis, 2)
       Case "280232"
             If Not (Trim(pdFecha) <> "\  \" Or Trim(pdFechaF) <> "\  \") Then
                   If Not (IsDate(pdFecha) = False Or IsDate(pdFechaF) = False) Then
                        MsgBox "INGRESE FECHA VALIDA", vbOKOnly + vbInformation, "AVISO"
                        Exit Sub
                   End If
             End If
             If Trim(psCodAge) = "" Then
                    MsgBox "Seleccione una Agencia de la Lista", vbOKOnly + vbInformation, "AVISO"
                    Exit Sub
             End If
                
              If psResp = "" Then
                MsgBox "No existe Informacion a mostrar", vbOKOnly, "AVISO"
              End If
              Exit Sub
                      
       Case "280851"
              Set rsTemp = orep.GetRepMejoresClientesN(CLng(pnMonto), pnTipoCambio, psCodAge, gdFecha, 1)
       Case "280852"
              Set rsTemp = orep.GetRepMejoresClientesN(CLng(pnMonto), pnTipoCambio, psCodAge, gdFecha, 2)
       Case "280853"
              Set rsTemp = orep.GetRepMejoresClientesN(CLng(pnMonto), pnTipoCambio, psCodAge, gdFecha, 3, "232")
       Case "280854"
              Set rsTemp = orep.GetRepMejoresClientesN(CLng(pnMonto), pnTipoCambio, psCodAge, gdFecha, 3, "233")
       Case "280855"
              Set rsTemp = orep.GetRepMejoresClientesN(CLng(pnMonto), pnTipoCambio, psCodAge, gdFecha, 3, "234")
       Case "280856"
              Set rsTemp = orep.GetRepMejoresClientesN(CLng(pnMonto), pnTipoCambio, psCodAge, gdFecha, 4, "232")
       Case "280857"
              Set rsTemp = orep.GetRepMejoresClientesN(CLng(pnMonto), pnTipoCambio, psCodAge, gdFecha, 4, "233")
End Select


If rsTemp.EOF Then
    MsgBox "No se encontro información para este reporte", vbOKOnly + vbInformation, "Aviso"
    Exit Sub
End If

Dim lsArchivoN As String, lbLibroOpen As Boolean
Dim sRep As String



   lsArchivoN = App.path & "\Spooler\Rep" & lsRep & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xls"
   
   'OleExcel.Class = "ExcelWorkSheet"
   lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
   If lbLibroOpen Then
            Set xlHoja1 = xlLibro.Worksheets(1)
            ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
            
                Select Case lsRep
                
                Case "280708", "280852"
                    sRep = " (CON IFIS)"
                    
                Case "280709", "280851"
                    sRep = " (SIN IFIS)"
                Case "280853"
                    sRep = " (SIN IFIS) - AHORROS"
                Case "280854"
                    sRep = " (SIN IFIS) - PLAZO FIJO"
                Case "280855"
                    sRep = " (SIN IFIS) - CTS"
                Case "280856"
                    sRep = " (CON IFIS) - AHORROS"
                Case "280857"
                    sRep = " (CON IFIS) - PLAZO FIJO"
                End Select
                       
            
            nFila = 1
            
            xlHoja1.Cells(nFila, 1) = gsNomCmac
            nFila = 2
            xlHoja1.Cells(nFila, 1) = gsNomAge
            xlHoja1.Range("F2:H2").MergeCells = True
            xlHoja1.Cells(nFila, 6) = Format(CDate(pdFecha), "Long Date")
             
             prgBar.value = 2
                 
             
            nFila = 3
            xlHoja1.Cells(nFila, 1) = "REPORTE DE LOS " & CStr(CLng(pnMonto)) & " MEJORES CLIENTES DE LA CAJA " & sRep
             
            
            xlHoja1.Range("A1:M5").Font.Bold = True
            
            xlHoja1.Range("A3:M3").MergeCells = True
            xlHoja1.Range("A3:A3").HorizontalAlignment = xlCenter
            xlHoja1.Range("A5:M5").HorizontalAlignment = xlCenter
 
            'xlHoja1.Range("A5:H5").AutoFilter
            
            nFila = 5
            
            xlHoja1.Cells(nFila, 1) = "ITEM"
            xlHoja1.Cells(nFila, 2) = "CODIGO"
            xlHoja1.Cells(nFila, 3) = "NOMBRE"
            xlHoja1.Cells(nFila, 4) = "DIRECCION"
            xlHoja1.Cells(nFila, 5) = "SALDO"
            xlHoja1.Cells(nFila, 6) = "FONO"
            xlHoja1.Cells(nFila, 7) = "FEC. NAC."
            xlHoja1.Cells(nFila, 8) = "ZONA"
            
            i = 0
            While Not rsTemp.EOF
                nFila = nFila + 1
                
                prgBar.value = ((i) / rsTemp.RecordCount) * 100
                
                i = i + 1

                
                xlHoja1.Cells(nFila, 1) = Format(i, "0000")
                xlHoja1.Cells(nFila, 2) = rsTemp!cCodPers
                xlHoja1.Cells(nFila, 3) = rsTemp!cNomPers
                xlHoja1.Cells(nFila, 4) = rsTemp!cDirPers
                xlHoja1.Cells(nFila, 5) = Format(rsTemp!nSaldo, "#,##0.00")
                xlHoja1.Cells(nFila, 6) = rsTemp!cTelPers & ""
                xlHoja1.Cells(nFila, 7) = Format(rsTemp!dFecNac, "dd/mm/yyyy hh:mm:ss")
                xlHoja1.Cells(nFila, 8) = rsTemp!Zona
                                            
                rsTemp.MoveNext
                
            Wend
            
           ' xlHoja1.Columns.AutoFit
            
            xlHoja1.Cells.Select
            xlHoja1.Cells.Font.Name = "Arial"
            xlHoja1.Cells.Font.Size = 9
            xlHoja1.Cells.EntireColumn.AutoFit
           
        
                
            'Cierro...
            'OleExcel.Class = "ExcelWorkSheet"
            ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
'            OleExcel.SourceDoc = lsArchivoN
'            OleExcel.Verb = 1
'            OleExcel.Action = 1
'            OleExcel.DoVerb -1
            
   End If
   
   Set rsTemp = Nothing
   


End Sub


Public Function ImpRepHavDevBoveda(ByVal psFecIni As Date, ByVal psFecFin As Date, ByVal psNomCmac As String, ByVal psNomAge As String, ByVal psCodAge As String) As String
  Dim cmovnro As String, rsTemp As New ADODB.Recordset
  Dim orep As cOMNCaptaGenerales.NCOMCaptaReportes
  Dim xlAplicacion As Excel.Application
  Dim xlLibro As Excel.Workbook
  Dim xlHoja1 As Excel.Worksheet
  Dim nFila As Long, i As Long
  Dim MONHAB  As Double, MONDEV As Double
  Dim NUMHAB  As Double, NUMDEV As Double
  Dim lsArchivoN As String, lbLibroOpen As Boolean
  

  MONHAB = 0
  MONDEV = 0
  NUMHAB = 0
  NUMDEV = 0
  
  
  Set orep = New cOMNCaptaGenerales.NCOMCaptaReportes
  
    Set rsTemp = orep.RepHavDevBovedaRango(Format(psFecIni, "yyyymmdd"), Format(psFecFin, "yyyymmdd"), psCodAge)
    
   
  If rsTemp.EOF Then
    MsgBox "No se encontro información para este reporte", vbOKOnly + vbInformation, "Aviso"
    Exit Function
  End If
   
  lsArchivoN = App.path & "\Spooler\RepHABDEVBOV" & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xls"
   
  'OLEEXCEL.Class = "ExcelWorkSheet"
  lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
  If lbLibroOpen Then
            Set xlHoja1 = xlLibro.Worksheets(1)
            ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
            
                                  
            
            nFila = 1
            
            xlHoja1.Cells(nFila, 1) = gsNomCmac
            nFila = 2
            xlHoja1.Cells(nFila, 1) = gsNomAge
            xlHoja1.Range("F2:H2").MergeCells = True
            xlHoja1.Cells(nFila, 6) = Format(gdFecSis, "Long Date")
             
'             prgBar.value = 2
                 
             
            nFila = 3
            xlHoja1.Cells(nFila, 1) = "REPORTE DE HABILITACIONES Y DEVOLUCIONES PARA BOVEDA " & psNomAge & " DEL" & Format(psFecIni, "dd/MM/yyyy") & IIf(psFecIni <> psFecFin, " AL" & Format(psFecFin, "dd/MM/yyyy"), "")
             
            
            xlHoja1.Range("A1:M3").Font.Bold = True
            
            xlHoja1.Range("A3:M3").MergeCells = True
            xlHoja1.Range("A3:A3").HorizontalAlignment = xlCenter
            
 
            'xlHoja1.Range("A5:H5").AutoFilter
            
            nFila = 5
            
                nFila = nFila + 1
                
            xlHoja1.Range("A" & nFila & ":E" & nFila).Font.Bold = True
            xlHoja1.Range("A" & nFila & ":E" & nFila).MergeCells = True
            xlHoja1.Range("A" & nFila & ":E" & nFila).HorizontalAlignment = xlCenter
            xlHoja1.Cells(nFila, 1) = "HABILITACIONES"
            
              nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "ITEM"
            xlHoja1.Cells(nFila, 2) = "MONEDA"
            xlHoja1.Cells(nFila, 3) = "IMPORTE"
            xlHoja1.Cells(nFila, 4) = "USUARIO"
            xlHoja1.Cells(nFila, 5) = "NOMBRE USUARIO"
            xlHoja1.Cells(nFila, 6) = "FECHA"
            xlHoja1.Cells(nFila, 7) = "HORA"
            
            i = 0
            While Not rsTemp.EOF
            
               If rsTemp.Fields("COPECOD") <> 901017 Then
                  GoTo Men
               End If
               
               

                nFila = nFila + 1
                
                
                
                i = i + 1

                
                xlHoja1.Cells(nFila, 1) = Format(i, "0000")
                xlHoja1.Cells(nFila, 2) = rsTemp!nmoneda
                xlHoja1.Cells(nFila, 3) = Format(rsTemp!NMOVIMPORTE, "#0.00")
                xlHoja1.Cells(nFila, 4) = rsTemp!CUSUDEST
                xlHoja1.Cells(nFila, 5) = rsTemp!NOMBRE
                xlHoja1.Cells(nFila, 6) = Format(CDate(Mid(rsTemp!cmovnro, 5, 2) & "-" & Mid(rsTemp!cmovnro, 7, 2) & "-" & Left(rsTemp!cmovnro, 4)), "dd/MM/yyyy")
                xlHoja1.Cells(nFila, 7) = Mid(rsTemp!cmovnro, 9, 2) & ":" & Mid(rsTemp!cmovnro, 11, 2) & ":" & Mid(rsTemp!cmovnro, 13, 2)
                
                                            
                MONHAB = MONHAB + rsTemp!NMOVIMPORTE
                
                rsTemp.MoveNext
                
                
            Wend
           
            
                          
Men:

            NUMHAB = i
            
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "TOTAL: " & CStr(NUMHAB)
            xlHoja1.Cells(nFila, 3) = Format(MONHAB, "#0.00")


            nFila = nFila + 2
                
            xlHoja1.Range("A" & nFila & ":E" & nFila).Font.Bold = True
            xlHoja1.Range("A" & nFila & ":E" & nFila).MergeCells = True
            xlHoja1.Range("A" & nFila & ":E" & nFila).HorizontalAlignment = xlCenter
            
            xlHoja1.Cells(nFila, 1) = "DEVOLUCIONES"
            
                
                nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "ITEM"
            xlHoja1.Cells(nFila, 2) = "MONEDA"
            xlHoja1.Cells(nFila, 3) = "IMPORTE"
            xlHoja1.Cells(nFila, 4) = "USUARIO"
            xlHoja1.Cells(nFila, 5) = "NOMBRE USUARIO"
            xlHoja1.Cells(nFila, 6) = "FECHA"
            xlHoja1.Cells(nFila, 7) = "HORA"
            
            i = 0
            While Not rsTemp.EOF


                nFila = nFila + 1
                
                
                
                i = i + 1
                
                xlHoja1.Cells(nFila, 1) = Format(i, "0000")
                xlHoja1.Cells(nFila, 2) = rsTemp!nmoneda
                xlHoja1.Cells(nFila, 3) = Format(rsTemp!NMOVIMPORTE, "#0.00")
                xlHoja1.Cells(nFila, 4) = rsTemp!CUSUDEST
                xlHoja1.Cells(nFila, 5) = rsTemp!NOMBRE
                xlHoja1.Cells(nFila, 6) = Format(CDate(Mid(rsTemp!cmovnro, 5, 2) & "-" & Mid(rsTemp!cmovnro, 7, 2) & "-" & Left(rsTemp!cmovnro, 4)), "dd/MM/yyyy")
                xlHoja1.Cells(nFila, 7) = Mid(rsTemp!cmovnro, 9, 2) & ":" & Mid(rsTemp!cmovnro, 11, 2) & ":" & Mid(rsTemp!cmovnro, 13, 2)
                                            
                MONDEV = MONDEV + rsTemp!NMOVIMPORTE
                
                rsTemp.MoveNext
                
            Wend
            NUMDEV = i
                       
                                              
            
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = "TOTAL: " & CStr(NUMDEV)
            xlHoja1.Cells(nFila, 3) = Format(MONDEV, "#0.00")
            
            Set rsTemp = New Recordset
            
            Set rsTemp = orep.REPBOVSALDOS(psCodAge, Format(psFecIni, "yyyymmdd"), Format(psFecFin, "yyyymmdd"))
            
               nFila = nFila + 1
               
             xlHoja1.Range("A" & nFila & ":E" & nFila).Font.Bold = True
             xlHoja1.Range("A" & nFila & ":E" & nFila).MergeCells = True
             xlHoja1.Range("A" & nFila & ":E" & nFila).HorizontalAlignment = xlCenter
           
                xlHoja1.Cells(nFila, 1) = "SALDOS FINALES"
                
                nFila = nFila + 1
                xlHoja1.Cells(nFila, 1) = "USUARIO"
                xlHoja1.Cells(nFila, 2) = "NOMBRE USUARIO"
                xlHoja1.Cells(nFila, 3) = "MONTO S/."
                xlHoja1.Cells(nFila, 4) = "MONTO U$."
                xlHoja1.Cells(nFila, 5) = "FECHA"
            
            While Not rsTemp.EOF
                            
                nFila = nFila + 1
                
                xlHoja1.Cells(nFila, 1) = rsTemp!Cuser
                xlHoja1.Cells(nFila, 2) = rsTemp!cPersNombre
                xlHoja1.Cells(nFila, 3) = rsTemp!solesmonto
                xlHoja1.Cells(nFila, 4) = rsTemp!dolaresmonto
                xlHoja1.Cells(nFila, 5) = Format(CDate(Mid(rsTemp!dFecha, 5, 2) & "/" & Right(rsTemp!dFecha, 2) & "/" & Left(rsTemp!dFecha, 4)), "dd/MM/yyyy")
                               
                    rsTemp.MoveNext
            Wend
            
                       
                        
            xlHoja1.Columns.AutoFit
            
            xlHoja1.Cells.Select
            xlHoja1.Cells.Font.Name = "Arial"
            xlHoja1.Cells.Font.Size = 9
            xlHoja1.Cells.EntireColumn.AutoFit
                   
                
            'Cierro...
'            OLEEXCEL.Class = "ExcelWorkSheet"
            ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
'            OLEEXCEL.SourceDoc = lsArchivoN
'            OLEEXCEL.Verb = 1
'            OLEEXCEL.Action = 1
'            OLEEXCEL.DoVerb -1
            
   End If
   
   ImpRepHavDevBoveda = "GENERADO"
      
   
   Set rsTemp = Nothing
   
  
End Function


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


