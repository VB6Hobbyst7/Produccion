Attribute VB_Name = "gRepCajaSaldoTiempoReal"
Option Explicit

Public Sub InsRow(ByRef MSFlex As MSHFlexGrid, Index As Integer)
If Index = 1 And MSFlex.Rows > 1 Then
   MSFlex.RowHeight(1) = 260
Else
   MSFlex.AddItem ""
   MSFlex.RowHeight(Index) = 260
End If
End Sub

'Public Function FNumero(vExpNumStr As Variant, Optional vNroDecimales As Integer) As String
'Dim nDec As Integer
'If Not IsNull(vExpNumStr) Then
'   nDec = IIf(vNroDecimales = 0, 2, vNroDecimales)
'   If InStr(",", vExpNumStr) = 0 Then
'      FNumero = Format(vExpNumStr, "###,###,##0." + String(nDec, "0"))
'   Else
'      FNumero = Format(Val(vExpNumStr), "###,###,##0." + String(nDec, "0"))
'   End If
'Else
'   FNumero = ""
'End If
'End Function
'
'
'
'Public Function DigNumEnt(KeyAscii As Integer, Optional vOtrosChar As String) As Integer
'Dim nPos As Integer
'nPos = InStr("0123456789" + Trim(vOtrosChar), Chr(KeyAscii))
'If nPos > 0 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Then
'   DigNumEnt = KeyAscii
'Else
'   Beep
'   DigNumEnt = 0
'End If
'End Function

Public Sub SelTexto(vCajaTexto As Control)
vCajaTexto.SelStart = 0
vCajaTexto.SelLength = Len(Trim(vCajaTexto))
End Sub

Public Function DigFecha(TextBox As Control, KeyAscii As Integer) As Integer
Dim nPos As Integer
nPos = InStr("0123456789", Chr(KeyAscii))
If nPos > 0 Then
   If Len(Trim(TextBox)) = 2 Or Len(Trim(TextBox)) = 5 Then
      TextBox = TextBox & "/"
      TextBox.SelStart = Len(TextBox)
   End If
   DigFecha = KeyAscii
ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Then
   DigFecha = KeyAscii
Else
   Beep
   DigFecha = 0
End If
End Function

'Public Function DigNumDec(CTRLTextBox As TextBox, KeyAscii As Integer) As Integer
'Dim nPos As Integer
'nPos = InStr(".0123456789", Chr(KeyAscii))
'If nPos > 0 Then
'   If Chr(KeyAscii) = "." And InStr(CTRLTextBox, Chr(KeyAscii)) > 0 Then
'      Beep
'      DigNumDec = 0
'   Else
'      DigNumDec = KeyAscii
'   End If
'ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Then
'   DigNumDec = KeyAscii
'Else
'   Beep
'   DigNumDec = 0
'End If
'End Function


Public Sub ImprimirSaldoTiempoReal()
On Error GoTo ImprimirSaldoTiempoRealErr
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1(100) As Excel.Worksheet
    Dim lbExisteHoja  As Boolean, liLineas As Integer, nTotalSoles As Currency, nTotalDolares As Currency
    Dim oCon As DConecta, sSql As String, rs As ADODB.Recordset, sSaldos As String
    Dim nSaldoSoles As Currency, nSaldoDolares As Currency, nPos As Integer, nNroHoja As Integer
    Set oCon = New DConecta
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    
    nNroHoja = 1
    ExcelAddHoja "Item " & nNroHoja, xlLibro, xlHoja1(nNroHoja), False
    
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(1, 1), xlHoja1(nNroHoja).Cells(1, 1)).ColumnWidth = 3
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(2, 2)).Font.Bold = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(2, 5)).Merge True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(4, 2)).Font.Bold = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 5)).Font.Bold = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 5)).Merge True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(7, 2), xlHoja1(nNroHoja).Cells(7, 5)).Font.Bold = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(7, 2), xlHoja1(nNroHoja).Cells(7, 5)).HorizontalAlignment = xlCenter
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 5)).HorizontalAlignment = xlCenter

    
    Screen.MousePointer = 11
    xlHoja1(nNroHoja).Cells(2, 2) = "Caja Municipal de Ahorros y Creditos de Trujillo SA"
    xlHoja1(nNroHoja).Cells(4, 2) = "Saldo de Caja en Tiempo Real"
    xlHoja1(nNroHoja).Cells(7, 2) = "#"
    xlHoja1(nNroHoja).Cells(7, 3) = "Agencia"
    xlHoja1(nNroHoja).Cells(7, 4) = "Saldo Soles"
    xlHoja1(nNroHoja).Cells(7, 5) = "Saldo Dolares"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(7, 2), xlHoja1(nNroHoja).Cells(7, 5)).Cells.Interior.Color = &HC0FFFF
    
    If oCon.AbreConexion Then
        sSql = "select cAgeCod,cAgeDescripcion from Agencias"
        Set rs = oCon.CargaRecordSet(sSql)
        oCon.CierraConexion
        liLineas = 9
    End If
    nTotalSoles = 0:       nTotalDolares = 0
    Do While Not rs.EOF
        If rs!cAgecod <> "04" And rs!cAgecod <> "11" Then
            nSaldoSoles = 0:    nSaldoDolares = 0:     sSaldos = ""
            sSaldos = SaldoAgencia(rs!cAgecod)
            If sSaldos <> "" Then
                xlHoja1(nNroHoja).Cells(liLineas, 2) = rs!cAgecod
                xlHoja1(nNroHoja).Cells(liLineas, 3) = rs!cAgeDescripcion
                nPos = InStr(1, sSaldos, "D")
                nSaldoSoles = Mid(sSaldos, 2, nPos - 2)
                nSaldoDolares = Mid(sSaldos, nPos + 1)
                xlHoja1(nNroHoja).Cells(liLineas, 4) = nSaldoSoles
                xlHoja1(nNroHoja).Cells(liLineas, 5) = nSaldoDolares
                nTotalSoles = nTotalSoles + nSaldoSoles
                nTotalDolares = nTotalDolares + nSaldoDolares
                liLineas = liLineas + 1
            End If
        End If
        rs.MoveNext
    Loop
    
    liLineas = liLineas + 1
        
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(liLineas, 2), xlHoja1(nNroHoja).Cells(liLineas, 5)).Font.Bold = True
    xlHoja1(nNroHoja).Cells(liLineas, 3) = "Total"
    xlHoja1(nNroHoja).Cells(liLineas, 4) = nTotalSoles
    xlHoja1(nNroHoja).Cells(liLineas, 5) = nTotalDolares
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(liLineas, 2), xlHoja1(nNroHoja).Cells(liLineas, 5)).Cells.Interior.Color = &HFCFFE1
    
    ExcelCuadro xlHoja1(nNroHoja), 2, 7, 5, liLineas, True, True
    
    Screen.MousePointer = 0
    '******************************************************
    
'        xlHoja1(nNroHoja).Cells.Select
    xlHoja1(nNroHoja).Cells.Font.Size = 8
    xlHoja1(nNroHoja).Cells.EntireColumn.AutoFit
    xlHoja1(nNroHoja).Cells.NumberFormat = "###,###,##0.00"

    'Libera los objetos.
    Set xlHoja1(nNroHoja) = Nothing
    nNroHoja = nNroHoja + 1
    
    xlLibro.SaveAs App.path & "\spooler\SaldoTiempoReal_" & Format(gdFecSis, "ddmmyyyy")
    Set xlLibro = Nothing
    xlAplicacion.Application.Visible = True
    xlAplicacion.Windows(1).Visible = True
    Exit Sub
ImprimirSaldoTiempoRealErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Function SaldoAgencia(ByVal pcAgeCod As String) As String
On Error GoTo SaldoAgenciaErr
    Dim oCon As DConecta, rs As ADODB.Recordset, sSql As String, nSumaSoles As Currency, nSumaDolares As Currency
    Set oCon = New DConecta
    If oCon.AbreConexion Then
'        sSQL = "[" & gsCentralImg & "].DBCmact" & pcAgeCod & ".dbo.CapSobrantesFaltantes "
        sSql = "[128.107.2.3].DBCmact" & pcAgeCod & ".dbo.CapSobrantesFaltantes "
        'sSQL = " CapSobrantesFaltantes "
        Set rs = oCon.CargaRecordSet(sSql)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    nSumaDolares = 0: nSumaSoles = 0
    Do While Not rs.EOF
        nSumaSoles = nSumaSoles + rs!MTSE
        nSumaDolares = nSumaDolares + rs!MTDE
        rs.MoveNext
    Loop
    SaldoAgencia = "S" & nSumaSoles & "D" & nSumaDolares
    rs.Close
    Set rs = Nothing
    Exit Function
SaldoAgenciaErr:
    Set rs = Nothing
    'MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function


'*********************************************************************************
'************************Reporte Adeudados Mi Vivienda****************************
'*********************************************************************************

Public Sub ImprimirAdeudadosMiVivienda(ByVal pcFecha As String, ByVal pnMoneda As Integer)
On Error GoTo ImprimirSaldoTiempoRealErr
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1(100) As Excel.Worksheet
    Dim lbExisteHoja  As Boolean, liLineas As Integer, nTotalSoles As Currency, nTotalDolares As Currency
    Dim oCon As DConecta, sSql As String, rs As ADODB.Recordset, sSaldos As String
    Dim nSaldoSoles As Currency, nSaldoDolares As Currency, nPos As Integer, nNroHoja As Integer
    Set oCon = New DConecta
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    
    nNroHoja = 1
    ExcelAddHoja "Item " & nNroHoja, xlLibro, xlHoja1(nNroHoja), False
    
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(1, 1), xlHoja1(nNroHoja).Cells(1, 1)).ColumnWidth = 3
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(2, 2)).Font.Bold = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(2, 12)).Merge True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(4, 2)).Font.Bold = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 12)).Font.Bold = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 12)).Merge True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 2), xlHoja1(nNroHoja).Cells(7, 13)).Font.Bold = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 2), xlHoja1(nNroHoja).Cells(7, 13)).HorizontalAlignment = xlCenter
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 12)).HorizontalAlignment = xlCenter

    
    Screen.MousePointer = 11
    xlHoja1(nNroHoja).Cells(2, 2) = "Caja Municipal de Ahorros y Creditos de Trujillo SA"
    xlHoja1(nNroHoja).Cells(4, 2) = "Reporte de Adeudados Mi Vivienda"
    
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 2), xlHoja1(nNroHoja).Cells(7, 2)).Merge False
    xlHoja1(nNroHoja).Cells(6, 2) = "Codigo"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 3), xlHoja1(nNroHoja).Cells(7, 3)).Merge False
    xlHoja1(nNroHoja).Cells(6, 3) = "Descripcion"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 4), xlHoja1(nNroHoja).Cells(6, 5)).Merge True
    xlHoja1(nNroHoja).Cells(6, 4) = "Fecha"
    xlHoja1(nNroHoja).Cells(7, 4) = "Apertura"
    xlHoja1(nNroHoja).Cells(7, 5) = "Vencimiento"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 6), xlHoja1(nNroHoja).Cells(7, 6)).Merge False
    xlHoja1(nNroHoja).Cells(6, 6) = "Credito"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 7), xlHoja1(nNroHoja).Cells(7, 7)).Merge False
    xlHoja1(nNroHoja).Cells(6, 7) = "Titular"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 8), xlHoja1(nNroHoja).Cells(6, 9)).Merge True
    xlHoja1(nNroHoja).Cells(6, 8) = "Desembolso"
    xlHoja1(nNroHoja).Cells(7, 8) = "Adeudo"
    xlHoja1(nNroHoja).Cells(7, 9) = "Credito"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 10), xlHoja1(nNroHoja).Cells(7, 10)).Merge False
    xlHoja1(nNroHoja).Cells(6, 10) = "Saldo Real"
    xlHoja1(nNroHoja).Cells(6, 11) = "Saldo Concesional"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 11), xlHoja1(nNroHoja).Cells(6, 12)).Merge True
    xlHoja1(nNroHoja).Cells(7, 11) = "Anterior"
    xlHoja1(nNroHoja).Cells(7, 12) = "Actual"
    xlHoja1(nNroHoja).Cells(6, 13) = "Saldo"
    xlHoja1(nNroHoja).Cells(7, 13) = "No Concesional"
    
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 2), xlHoja1(nNroHoja).Cells(7, 13)).Cells.Interior.Color = &HC0FFFF
    
    If oCon.AbreConexion Then
        sSql = "exec dbo.CGAdeudadosMiVivienda '" & Format(pcFecha, "yyyymmdd") & "', '" & pnMoneda & "'"
        Set rs = oCon.CargaRecordSet(sSql)
        oCon.CierraConexion
        liLineas = 9
    End If
    nTotalSoles = 0:       nTotalDolares = 0
    Do While Not rs.EOF
        xlHoja1(nNroHoja).Cells(liLineas, 2) = rs!cCtaIfCod
        xlHoja1(nNroHoja).Cells(liLineas, 3) = rs!cCtaIFDesc
        xlHoja1(nNroHoja).Cells(liLineas, 4) = rs!dCtaIFAper
        xlHoja1(nNroHoja).Cells(liLineas, 5) = rs!FechaVen
        xlHoja1(nNroHoja).Cells(liLineas, 6) = rs!Credito
        xlHoja1(nNroHoja).Cells(liLineas, 7) = rs!Titular
        xlHoja1(nNroHoja).Cells(liLineas, 8) = rs!nMontoPrestado
        xlHoja1(nNroHoja).Cells(liLineas, 9) = rs!nMontoPrestadoReal
        xlHoja1(nNroHoja).Cells(liLineas, 10) = rs!nSaldoNewReal
        xlHoja1(nNroHoja).Cells(liLineas, 11) = rs!SaldoConcecionalAnt
        xlHoja1(nNroHoja).Cells(liLineas, 12) = rs!SaldoConcecionalAct
        xlHoja1(nNroHoja).Cells(liLineas, 13) = rs!nSaldoNewReal - rs!SaldoConcecionalAct
        
        liLineas = liLineas + 1
        rs.MoveNext
    Loop
    
    liLineas = liLineas + 1
        
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(liLineas, 2), xlHoja1(nNroHoja).Cells(liLineas, 13)).Font.Bold = True
    xlHoja1(nNroHoja).Cells(liLineas, 3) = "Total"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(liLineas, 3), xlHoja1(nNroHoja).Cells(liLineas, 7)).Merge True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(liLineas, 2), xlHoja1(nNroHoja).Cells(liLineas, 13)).Cells.Interior.Color = &HFCFFE1
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(9, 4), xlHoja1(nNroHoja).Cells(liLineas, 5)).NumberFormat = "dd/mm/yyyy"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(9, 8), xlHoja1(nNroHoja).Cells(liLineas, 13)).NumberFormat = "###,###,##0.00"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(liLineas, 8), xlHoja1(nNroHoja).Cells(liLineas, 8)).Formula = "=Sum(H9:H" & liLineas - 2 & ")"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(liLineas, 9), xlHoja1(nNroHoja).Cells(liLineas, 9)).Formula = "=Sum(I9:I" & liLineas - 2 & ")"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(liLineas, 10), xlHoja1(nNroHoja).Cells(liLineas, 10)).Formula = "=Sum(J9:J" & liLineas - 2 & ")"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(liLineas, 11), xlHoja1(nNroHoja).Cells(liLineas, 11)).Formula = "=Sum(K9:K" & liLineas - 2 & ")"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(liLineas, 12), xlHoja1(nNroHoja).Cells(liLineas, 12)).Formula = "=Sum(L9:L" & liLineas - 2 & ")"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(liLineas, 13), xlHoja1(nNroHoja).Cells(liLineas, 13)).Formula = "=Sum(M9:M" & liLineas - 2 & ")"
    
    ExcelCuadro xlHoja1(nNroHoja), 2, 6, 13, liLineas, True, True
    
    Screen.MousePointer = 0
    '******************************************************
    
'        xlHoja1(nNroHoja).Cells.Select
    xlHoja1(nNroHoja).Cells.Font.Size = 8
    xlHoja1(nNroHoja).Cells.EntireColumn.AutoFit
    'xlHoja1(nNroHoja).Cells.NumberFormat = "###,###,##0.00"

    'Libera los objetos.
    Set xlHoja1(nNroHoja) = Nothing
    nNroHoja = nNroHoja + 1
    
    xlLibro.SaveAs App.path & "\spooler\AdeudadoMiVivienda_" & Format(pcFecha, "ddmmyyyy") & "_" & pnMoneda
    Set xlLibro = Nothing
    xlAplicacion.Application.Visible = True
    xlAplicacion.Windows(1).Visible = True
    Exit Sub
ImprimirSaldoTiempoRealErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

'*********************************************************************************

Public Function DigPeriodo(TextBox As Control, KeyAscii As Integer) As Integer
Dim nPos As Integer
nPos = InStr("0123456789", Chr(KeyAscii))
If nPos > 0 Then
   If Len(Trim(TextBox)) = 2 Then
      TextBox = TextBox & "/"
      TextBox.SelStart = Len(TextBox)
   End If
   DigPeriodo = KeyAscii
ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Then
   DigPeriodo = KeyAscii
Else
   DigPeriodo = 0
End If
End Function
