Attribute VB_Name = "gFunACL"

Public Sub nRepo179104_ColocacionFoncodes(ByVal psServConsol As String, ByVal psCmact As String, ByVal pdFechaCierre As Date)
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim fs As Scripting.FileSystemObject

Dim sql As String
Dim rs As New ADODB.Recordset

Dim lsArchivo As String
Dim nFil As Integer
Dim nCol As Integer
Dim lsCodigo As String
Dim Titulo As String
Dim i As Integer
Dim J As Integer

Dim Mes(1 To 12) As String
Dim nMes As Integer
Dim sFechaIni As String

Dim oCapRCD As COMNCaptaGenerales.NCOMRCDReportes
Set oCapRCD = New COMNCaptaGenerales.NCOMRCDReportes

Mes(1) = "ENE"
Mes(2) = "FEB"
Mes(3) = "MAR"
Mes(4) = "ABR"
Mes(5) = "MAY"
Mes(6) = "JUN"
Mes(7) = "JUL"
Mes(8) = "AGO"
Mes(9) = "SET"
Mes(10) = "OCT"
Mes(11) = "NOV"
Mes(12) = "DIC"
nMes = Month(pdFechaCierre)
Dim s As Integer
sFechaIni = DateAdd("m", -5, Format(pdFechaCierre, "mm/dd/yyyy"))

'RaiseEvent ShowProgress
'On Error Resume Next
lsArchivo = "ReporteMesFONCODES" & Format(pdFechaCierre, "YYYYMMDD") & ".xls"
Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

Set xlHoja1 = xlLibro.Worksheets.Add

xlHoja1.Range("A1").Cells.Font.Bold = True
xlHoja1.Range("A1").Cells.Font.Size = 10
xlHoja1.Cells(1, 1) = "INFORME MENSUAL DE COLOCACIONES DEL FONDO NOTATORIO"
xlHoja1.Range("A3:A4").Cells.Font.Size = 10
xlHoja1.Range("A3:A4").Cells.Font.Bold = True
xlHoja1.Cells(3, 1) = "NOMBRE DE LA INSTITUCION: "
xlHoja1.Cells(4, 1) = "MES DEL INFORME"

xlHoja1.Range("C6:H6").Cells.Font.Size = 8
xlHoja1.Range("C6:H6").Cells.Font.Bold = True
xlHoja1.Cells(6, 3) = "'" & Mes(Val(Month(sFechaIni))) & "-" & Format(sFechaIni, "YY")
xlHoja1.Cells(6, 4) = "'" & Mes(Val(Month(DateAdd("m", 1, sFechaIni)))) & " - " & Format(DateAdd("m", 1, sFechaIni), "YY")
xlHoja1.Cells(6, 5) = "'" & Mes(Val(Month(DateAdd("m", 2, sFechaIni)))) & " - " & Format(DateAdd("m", 2, sFechaIni), "YY")
xlHoja1.Cells(6, 6) = "'" & Mes(Val(Month(DateAdd("m", 3, sFechaIni)))) & " - " & Format(DateAdd("m", 3, sFechaIni), "YY")
xlHoja1.Cells(6, 7) = "'" & Mes(Val(Month(DateAdd("m", 4, sFechaIni)))) & " - " & Format(DateAdd("m", 4, sFechaIni), "YY")
xlHoja1.Cells(6, 8) = "'" & Mes(Val(Month(pdFechaCierre))) & " - " & Format(pdFechaCierre, "YY")


xlHoja1.Cells(6, 1) = "N°"
xlHoja1.Cells(6, 2) = "DESCRIPCION"
xlHoja1.Range("A6:H6").Cells.BorderAround xlContinuous

xlHoja1.Range("A8").Cells.Font.Bold = True
xlHoja1.Range("A8").Cells.Font.Size = 10
xlHoja1.Cells(8, 1) = "1.  SOLICITUDES PRESENTADAS"

xlHoja1.Cells(9, 1) = "1.1"
xlHoja1.Cells(9, 2) = "NUMERO"
xlHoja1.Cells(10, 1) = "1.20"
xlHoja1.Cells(10, 2) = "MONTO TOTAL NS/."

xlHoja1.Range("A11").Cells.Font.Bold = True
xlHoja1.Range("A11").Cells.Font.Size = 10
xlHoja1.Cells(11, 1) = "2.  CREDITOS APROBADOS"

xlHoja1.Cells(12, 1) = "2.1"
xlHoja1.Cells(12, 2) = "NUMERO"
xlHoja1.Cells(13, 1) = "2.20"
xlHoja1.Cells(13, 2) = "MONTO TOTAL NS/."

xlHoja1.Range("A14").Cells.Font.Bold = True
xlHoja1.Range("A14").Cells.Font.Size = 10
xlHoja1.Cells(14, 1) = "3.  CREDITOS DESEMBOLSADOS"

xlHoja1.Cells(15, 1) = "3.1"
xlHoja1.Cells(15, 2) = "NUMERO"
xlHoja1.Cells(16, 1) = "3.20"
xlHoja1.Cells(16, 2) = "MONTO TOTAL NS/."

xlHoja1.Range("A17:A18").Cells.Font.Bold = True
xlHoja1.Range("A17:A18").Cells.Font.Size = 10
xlHoja1.Cells(17, 1) = "4.  TASA INTERES PROMEDIO"
xlHoja1.Cells(18, 1) = "5.  PLAZOS"

xlHoja1.Cells(19, 1) = "5.1."
xlHoja1.Cells(20, 1) = "5.1.1"
xlHoja1.Cells(21, 1) = "5.1.2"
xlHoja1.Cells(22, 1) = "5.1.3"
xlHoja1.Cells(24, 1) = "5.2."
xlHoja1.Cells(25, 1) = "5.2.1"
xlHoja1.Cells(26, 1) = "5.2.2"
xlHoja1.Cells(27, 1) = "5.2.3"

xlHoja1.Cells(19, 2) = "NUMERO"
xlHoja1.Cells(20, 2) = "HASTA 6 MESES"
xlHoja1.Cells(21, 2) = "DE 6 A 12 MESES"
xlHoja1.Cells(22, 2) = "DE 12 A 24 MESES"
xlHoja1.Cells(23, 2) = "MAS DE 24 MESES"
xlHoja1.Cells(24, 2) = "MONTO TOTAL NS/."
xlHoja1.Cells(25, 2) = "HASTA 6 MESES"
xlHoja1.Cells(26, 2) = "DE 6 A 12 MESES"
xlHoja1.Cells(27, 2) = "DE 12 A 24 MESES"
xlHoja1.Cells(28, 2) = "MAS DE 24 MESES"

xlHoja1.Range("A29").Cells.Font.Bold = True
xlHoja1.Range("A29").Cells.Font.Size = 10
xlHoja1.Cells(29, 1) = "6.  COLOCACIONES POR SEXO (*)"

xlHoja1.Cells(30, 1) = "6.1."
xlHoja1.Cells(31, 1) = "6.1.1"
xlHoja1.Cells(32, 1) = "6.1.2"
xlHoja1.Cells(34, 1) = "6.2."
xlHoja1.Cells(35, 1) = "6.2.1"
xlHoja1.Cells(36, 1) = "6.2.2"

xlHoja1.Cells(30, 2) = "NUMERO"
xlHoja1.Cells(31, 2) = "HOMBRES"
xlHoja1.Cells(32, 2) = "MUJERES"
xlHoja1.Cells(33, 2) = "PERS JURIDICA"
xlHoja1.Cells(34, 2) = "MONTO TOTAL NS/."
xlHoja1.Cells(35, 2) = "HOMBRES"
xlHoja1.Cells(36, 2) = "MUJERES"
xlHoja1.Cells(37, 2) = "PERS JURIDICA"

xlHoja1.Range("A38").Cells.Font.Bold = True
xlHoja1.Range("A38").Cells.Font.Size = 10
xlHoja1.Cells(38, 1) = "7.  COLOCACIONES POR N° DE TRABAJADORES"

xlHoja1.Cells(39, 1) = "7.1."
xlHoja1.Cells(40, 1) = "7.1.1"
xlHoja1.Cells(41, 1) = "7.1.2"
xlHoja1.Cells(42, 1) = "7.2."
xlHoja1.Cells(43, 1) = "7.2.1"
xlHoja1.Cells(44, 1) = "7.2.2"

xlHoja1.Cells(39, 2) = "NUMERO DE TRABAJADORES POR EMPRESA"
xlHoja1.Cells(40, 2) = "MENOS DE 10 TRABAJADORES"
xlHoja1.Cells(41, 2) = "MAS DE 10 TRABAJADORES"
xlHoja1.Cells(42, 2) = "MONTO TOTAL NS/."
xlHoja1.Cells(43, 2) = "MENOS DE 10 TRABAJADORES"
xlHoja1.Cells(44, 2) = "MAS DE 10 TRABAJADORES"

xlHoja1.Range("A45").Cells.Font.Bold = True
xlHoja1.Range("A45").Cells.Font.Size = 10
xlHoja1.Cells(45, 1) = "8.  COLOCACIONES POR SECTOR"

xlHoja1.Cells(46, 1) = "8.1."
xlHoja1.Cells(47, 1) = "8.1.1"
xlHoja1.Cells(48, 1) = "8.1.2"
xlHoja1.Cells(49, 1) = "8.1.3"
xlHoja1.Cells(50, 1) = "8.1.4"
xlHoja1.Cells(51, 1) = "8.1.5"
xlHoja1.Cells(52, 1) = "8.2."
xlHoja1.Cells(53, 1) = "8.2.1"
xlHoja1.Cells(54, 1) = "8.2.2"
xlHoja1.Cells(55, 1) = "8.2.3"
xlHoja1.Cells(56, 1) = "8.2.4"
xlHoja1.Cells(57, 1) = "8.2.5"

xlHoja1.Cells(46, 2) = "NUMERO"
xlHoja1.Cells(47, 2) = "PRODUCCION"
xlHoja1.Cells(48, 2) = "SERVICIOS"
xlHoja1.Cells(49, 2) = "COMERCIO"
xlHoja1.Cells(50, 2) = "AGROPECUARIO"
xlHoja1.Cells(51, 2) = "EDUCACION"
xlHoja1.Cells(52, 2) = "MONTO TOTAL NS/."
xlHoja1.Cells(53, 2) = "PRODUCCION"
xlHoja1.Cells(54, 2) = "SERVICIOS"
xlHoja1.Cells(55, 2) = "COMERCIO"
xlHoja1.Cells(56, 2) = "AGROPECUARIO"
xlHoja1.Cells(57, 2) = "EDUCACION"

xlHoja1.Range("A58").Cells.Font.Bold = True
xlHoja1.Range("A58").Cells.Font.Size = 10
xlHoja1.Cells(58, 1) = "9.1  NUMERO"

xlHoja1.Cells(59, 1) = "9.1.1"
xlHoja1.Cells(60, 1) = "9.1.2"
xlHoja1.Cells(61, 1) = "9.1.3"
xlHoja1.Cells(62, 1) = "9.2."
xlHoja1.Cells(63, 1) = "9.2.1"
xlHoja1.Cells(64, 1) = "9.2.2"
xlHoja1.Cells(65, 1) = "9.2.3"

xlHoja1.Cells(59, 2) = "CAPITAL FIJO"
xlHoja1.Cells(60, 2) = "CAPITAL DE TRABAJO"
xlHoja1.Cells(61, 2) = "CAPITAL FIJO Y DE TRABAJO"
xlHoja1.Cells(62, 2) = "MONTO TOTAL NS/."
xlHoja1.Cells(63, 2) = "CAPITAL FIJO"
xlHoja1.Cells(64, 2) = "CAPITAL DE TRABAJO"
xlHoja1.Cells(65, 2) = "CAPITAL FIJO Y DE TRABAJO"

xlHoja1.Range("A66").Cells.Font.Bold = True
xlHoja1.Range("A66").Cells.Font.Size = 10
xlHoja1.Cells(66, 1) = "10  COLOCACIONES POR MONTO"

xlHoja1.Cells(67, 1) = "10.1"
xlHoja1.Cells(68, 1) = "10.1.1"
xlHoja1.Cells(69, 1) = "10.1.2"
xlHoja1.Cells(70, 1) = "10.1.3"
xlHoja1.Cells(71, 1) = "10.1.4"
xlHoja1.Cells(72, 1) = "10.20"
xlHoja1.Cells(73, 1) = "10.2.1"
xlHoja1.Cells(74, 1) = "10.2.2"
xlHoja1.Cells(75, 1) = "10.2.3"
xlHoja1.Cells(76, 1) = "10.2.4"

xlHoja1.Cells(67, 2) = "NUMERO"
xlHoja1.Cells(68, 2) = "HASTA 2,000 NS/."
xlHoja1.Cells(69, 2) = "DE 2,001 HASTA 5,000 NS/."
xlHoja1.Cells(70, 2) = "DE 5,001 HASTA 10,000 NS/."
xlHoja1.Cells(71, 2) = "MAS DE 10,000 NS/."
xlHoja1.Cells(72, 2) = "MONTO TOTAL NS/."
xlHoja1.Cells(73, 2) = "HASTA 2,000 NS/."
xlHoja1.Cells(74, 2) = "DE 2,001 HASTA 5,000 NS/."
xlHoja1.Cells(75, 2) = "DE 5,001 HASTA 10,000 NS/."
xlHoja1.Cells(76, 2) = "MAS DE 10,000 NS/."

xlHoja1.Range("A83").Cells.Font.Bold = True
xlHoja1.Range("A83").Cells.Font.Size = 8
xlHoja1.Cells(83, 1) = "(*) LAS COLOCACIONES SE REFIEREN AL NUMERO Y MONTO DE CREEDITOS DESEMBOLSADOS"

xlHoja1.Range("A1").ColumnWidth = 5
xlHoja1.Range("B1").ColumnWidth = 45
xlHoja1.Range("C1").ColumnWidth = 7
xlHoja1.Range("D1").ColumnWidth = 7
xlHoja1.Range("E1").ColumnWidth = 7
xlHoja1.Range("F1").ColumnWidth = 7
xlHoja1.Range("G1").ColumnWidth = 7
xlHoja1.Range("H1").ColumnWidth = 7

'querys
'Solicitudes Presentadas
'Funcion ObtenerSolicitudesPresentadas

Set rs = oCapRCD.ObtenerSolicitudesPresentadas(sFechaIni, pdFechaCierre)
While Not rs.EOF
    xlHoja1.Cells(9, i) = rs!Numero
    xlHoja1.Cells(10, i) = Format(rs!Monto, "0.00")
    i = i - 1
    rs.MoveNext
Wend
rs.Close

' CREDITOS APROBADOS
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerCreditosAprobados(sFechaIni, pdFechaCierre)
i = 8
While Not rs.EOF
    xlHoja1.Cells(12, i) = rs!Numero
    xlHoja1.Cells(13, i) = Format(rs!Monto, "0.00")
    i = i - 1
    rs.MoveNext
Wend
rs.Close

' CREDITOS DESEMBOLSADOS
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerCreditosDesembolsados(sFechaIni, pdFechaCierre)
i = 8
While Not rs.EOF
    xlHoja1.Cells(12, i) = rs!Numero
    xlHoja1.Cells(13, i) = Format(rs!Monto, "0.00")
    i = i - 1
    rs.MoveNext
Wend
rs.Close

' PLAZOS
Set rs = New ADODB.Recordset
For i = 0 To 5
    Set rs = oCapRCD.ObtenerPlazos(sFechaIni)
    J = 3
    If rs.EOF And rs.BOF Then
    Else
        While Not rs.EOF
            Select Case rs!Rango
                Case "01. Hasta 6 Meses"
                        xlHoja1.Cells(20, J) = rs!Num
                        xlHoja1.Cells(25, J) = rs!Desembolso
                Case "02. De 6 a 12 Meses"
                        xlHoja1.Cells(21, J) = rs!Num
                        xlHoja1.Cells(26, J) = rs!Desembolso
                Case "03. De 12 a 24 Meses"
                        xlHoja1.Cells(22, J) = rs!Num
                        xlHoja1.Cells(27, J) = rs!Desembolso
                Case "04. Mas de 24 Meses"
                        xlHoja1.Cells(23, J) = rs!Num
                        xlHoja1.Cells(28, J) = rs!Desembolso
            End Select
            rs.MoveNext
        Wend
    End If
    J = J + 1
Next i
rs.Close

'''''''''''''

'''''''''''''


'POR SEXO
Set rs = New ADODB.Recordset
For i = 0 To 5
    Set rs = oCapRCD.ObtenerPlazoporSexo(sFechaIni)
    J = 3
    If rs.EOF And rs.BOF Then
    Else
        While Not rs.EOF
            Select Case rs!cSexPers
                Case "M":
                    xlHoja1.Cells(32, J) = rs!Cuenta
                    xlHoja1.Cells(36, J) = rs!Desembolso
                Case "H":
                    xlHoja1.Cells(31, J) = rs!Cuenta
                    xlHoja1.Cells(35, J) = rs!Desembolso
                Case "J":
                    xlHoja1.Cells(33, J) = rs!Cuenta
                    xlHoja1.Cells(37, J) = rs!Desembolso
            End Select
            rs.MoveNext
        Wend
    End If
    J = J + 1
Next i
rs.Close

' COLOCACIONES POR NUMERO DE TRABAJADORES
Set rs = New ADODB.Recordset
For i = 0 To 5
    'ObtenerColocacionesporTrabajadores
    Set rs = oCapRCD.ObtenerColocacionesporTrabajador(sFechaIni)
    J = 3
    If rs.EOF And rs.BOF Then
    Else
        While Not rs.EOF
            If rs!NroTaba = "01. Hasta 10" Then
                xlHoja1.Cells(40, J) = rs!Num
                xlHoja1.Cells(41, J) = rs!Desembolso
            Else
                xlHoja1.Cells(43, J) = rs!Cuenta
                xlHoja1.Cells(44, J) = rs!Desembolso
            End If
            rs.MoveNext
        Wend
    End If
    J = J + 1
Next i
rs.Close

' COLOCACIONES POR MONTO

' ++++++++++++++++++++++
' ++++++++++++++++++++++

' Falta por Query Sector

' ++++++++++++++++++++++
' ++++++++++++++++++++++
Set rs = New ADODB.Recordset
For i = 0 To 5
    Set rs = oCapRCD.ObtenerColocacionesporMonto(sFechaIni)
    J = 3
    If rs.EOF And rs.BOF Then
    Else
        While Not rs.EOF
            Select Case rs!cDesCre
            Case "A":
                    xlHoja1.Cells(61, J) = rs!Cuenta
                    xlHoja1.Cells(65, J) = rs!Desembolso
            Case "B":
                    xlHoja1.Cells(60, J) = rs!Cuenta
                    xlHoja1.Cells(64, J) = rs!Desembolso
            Case "C":
                    xlHoja1.Cells(59, J) = rs!Cuenta
                    xlHoja1.Cells(63, J) = rs!Desembolso
            End Select
            J = J + 1
            rs.MoveNext
        Wend
    End If
Next i
rs.Close

'  COLOCACIONES POR MONTO
Set rs = New ADODB.Recordset
For i = 0 To 5
    Set rs = oCapRCD.ObtenerColocacinesporMonto1(sFechaIni)
    J = 3
    If rs.EOF And rs.BOF Then
    Else
    While Not rs.EOF
        Select Case rs!Rango
        Case "01. Hasta 2000 NS"
                    xlHoja1.Cells(68, J) = rs!Num
                    xlHoja1.Cells(73, J) = rs!Desembolso
        Case "02. De 2001 A 5000"
                    xlHoja1.Cells(69, J) = rs!Num
                    xlHoja1.Cells(74, J) = rs!Desembolso
        Case "03. De 5001 a 10000"
                    xlHoja1.Cells(70, J) = rs!Num
                    xlHoja1.Cells(75, J) = rs!Desembolso
        Case "04. Mas de 10000"
                    xlHoja1.Cells(71, J) = rs!Num
                    xlHoja1.Cells(76, J) = rs!Desembolso
        End Select
        rs.MoveNext
    Wend
    End If
    J = J + 1
Next i
rs.Close

' Suma de las operaciones
For i = Asc("C") To Asc("H")
    xlHoja1.Range(Chr(i) & "19").Formula = "=" & Chr(i) & "20 +" & Chr(i) & "21 +" & Chr(i) & "22"
    xlHoja1.Range(Chr(i) & "24").Formula = "=" & Chr(i) & "25 +" & Chr(i) & "26 +" & Chr(i) & "27 +" & Chr(i) & "28"
    xlHoja1.Range(Chr(i) & "30").Formula = "=" & Chr(i) & "31 +" & Chr(i) & "32"
    xlHoja1.Range(Chr(i) & "34").Formula = "=" & Chr(i) & "35 +" & Chr(i) & "37"
    xlHoja1.Range(Chr(i) & "39").Formula = "=" & Chr(i) & "40 +" & Chr(i) & "41"
    xlHoja1.Range(Chr(i) & "42").Formula = "=" & Chr(i) & "43 +" & Chr(i) & "44"
    xlHoja1.Range(Chr(i) & "46").Formula = "=" & Chr(i) & "47 +" & Chr(i) & "48 +" & Chr(i) & "49 +" & Chr(i) & "50 +" & Chr(i) & "51"
    xlHoja1.Range(Chr(i) & "52").Formula = "=" & Chr(i) & "53 +" & Chr(i) & "54 +" & Chr(i) & "55 +" & Chr(i) & "56 +" & Chr(i) & "57"
    xlHoja1.Range(Chr(i) & "59").Formula = "=" & Chr(i) & "60 +" & Chr(i) & "61"
    xlHoja1.Range(Chr(i) & "62").Formula = "=" & Chr(i) & "64 +" & Chr(i) & "65"
    xlHoja1.Range(Chr(i) & "67").Formula = "=" & Chr(i) & "68 +" & Chr(i) & "69 +" & Chr(i) & "70 +" & Chr(i) & "71"
    xlHoja1.Range(Chr(i) & "72").Formula = "=" & Chr(i) & "73 +" & Chr(i) & "74 +" & Chr(i) & "75 +" & Chr(i) & "76"
Next i

'------------- BORDES
xlHoja1.Range("C9:H10").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("C9:H10").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("C9:H10").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("C9:H10").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("C9:H10").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("C9:H10").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("C9:H10").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("C9:H10").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

xlHoja1.Range("C12:H13").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("C12:H13").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("C12:H13").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("C12:H13").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("C12:H13").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("C12:H13").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("C12:H13").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("C12:H13").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

xlHoja1.Range("C15:H16").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("C15:H16").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("C15:H16").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("C15:H16").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("C15:H16").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("C15:H16").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("C15:H16").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("C15:H16").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

xlHoja1.Range("C19:H28").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("C19:H28").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("C19:H28").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("C19:H28").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("C19:H28").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("C19:H28").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("C19:H28").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("C19:H28").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

xlHoja1.Range("C30:H37").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("C30:H37").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("C30:H37").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("C30:H37").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("C30:H37").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("C30:H37").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("C30:H37").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("C30:H37").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

xlHoja1.Range("C39:H44").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("C39:H44").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("C39:H44").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("C39:H44").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("C39:H44").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("C39:H44").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("C39:H44").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("C39:H44").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

xlHoja1.Range("C46:H57").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("C46:H57").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("C46:H57").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("C46:H57").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("C46:H57").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("C46:H57").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("C46:H57").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("C46:H57").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

xlHoja1.Range("C59:H65").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("C59:H65").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("C59:H65").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("C59:H65").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("C59:H65").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("C59:H65").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("C59:H65").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("C59:H65").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

xlHoja1.Range("C67:H76").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("C67:H76").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("C67:H76").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("C67:H76").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("C67:H76").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("C67:H76").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("C67:H76").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("C67:H76").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

Set oCapRCD = Nothing

xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'RaiseEvent Progress(20, 20)
'---------------->
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
'Set oTipCambio = Nothing
'-------------->

'RaiseEvent CloseProgress
MsgBox "Se ha Generado el Archivo " & lsArchivo & ".XLS Satisfactoriamente", vbInformation, "Aviso"
Exit Sub
ErrorExcel:
    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
    xlLibro.Close
    ' Cierra Microsoft Excel con el método Quit.
    xlAplicacion.Quit
    'Libera los objetos.
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing
End Sub

Public Sub nRepo179105_CarteraMensualFoncodes(ByVal psCmact As String, ByVal pdFechaCierre As Date)
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim fs As Scripting.FileSystemObject

Dim sql As String
Dim rs As New ADODB.Recordset

Dim lsArchivo As String
Dim nFil As Integer
Dim nCol As Integer
Dim lsCodigo As String
Dim Titulo As String
Dim i As Integer

Dim J As Integer

Dim Mes(1 To 12) As String
Dim nMes As Integer
Dim sFechaIni As String

Dim oCapRCD As COMNCaptaGenerales.NCOMRCDReportes
Set oCapRCD = New COMNCaptaGenerales.NCOMRCDReportes
Dim sValor As String

Mes(1) = "ENE"
Mes(2) = "FEB"
Mes(3) = "MAR"
Mes(4) = "ABR"
Mes(5) = "MAY"
Mes(6) = "JUN"
Mes(7) = "JUL"
Mes(8) = "AGO"
Mes(9) = "SET"
Mes(10) = "OCT"
Mes(11) = "NOV"
Mes(12) = "DIC"
nMes = Month(pdFechaCierre)
Dim s As Integer
sFechaIni = DateAdd("m", -5, pdFechaCierre)
'RaiseEvent ShowProgress
'On Error Resume Next

lsArchivo = "ReporteMens_CarteraFONCODES" & Format(pdFechaCierre, "YYYYMMDD") & ".xls"
Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

Set xlHoja1 = xlLibro.Worksheets.Add

xlHoja1.Range("A1").Cells.Font.Bold = True
xlHoja1.Range("A1").Cells.Font.Size = 10
xlHoja1.Cells(1, 1) = "INFORME MENSUAL DE SITUACION DE CARTERA DEL FONDO ROTATORIO"
xlHoja1.Range("A3").Cells.Font.Size = 10
xlHoja1.Range("A3").Cells.Font.Bold = True
xlHoja1.Cells(3, 1) = "NOMBRE DE LA INSTITUCION: "

xlHoja1.Cells(4, 1) = "N°"
xlHoja1.Cells(4, 2) = "DESCRIPCION"

xlHoja1.Cells(4, 3) = "'" & Mes(Val(Month(sFechaIni))) & " - " & Format(sFechaIni, "YY")
xlHoja1.Cells(4, 4) = "'" & Mes(Val(Month(DateAdd("m", 1, sFechaIni)))) & " - " & Format(DateAdd("m", 1, sFechaIni), "YY")
xlHoja1.Cells(4, 5) = "'" & Mes(Val(Month(DateAdd("m", 2, sFechaIni)))) & " - " & Format(DateAdd("m", 2, sFechaIni), "YY")
xlHoja1.Cells(4, 6) = "'" & Mes(Val(Month(DateAdd("m", 3, sFechaIni)))) & " - " & Format(DateAdd("m", 3, sFechaIni), "YY")
xlHoja1.Cells(4, 7) = "'" & Mes(Val(Month(DateAdd("m", 4, sFechaIni)))) & " - " & Format(DateAdd("m", 4, sFechaIni), "YY")
xlHoja1.Cells(4, 8) = "'" & Mes(Val(Month(pdFechaCierre))) & " - " & Format(pdFechaCierre, "YY")
xlHoja1.Range("C4:H4").Cells.Font.Bold = True
xlHoja1.Range("C4:H4").Cells.Font.Size = 9
xlHoja1.Range("C4:H4").HorizontalAlignment = xlCenter

'xlHoja1.Range("C4:H4").NumberFormat = "#,##0.00"

xlHoja1.Range("A5").Cells.Font.Bold = True
xlHoja1.Range("A5").Cells.Font.Size = 10
xlHoja1.Cells(5, 1) = "1.   CARTERA TOTAL"


xlHoja1.Cells(6, 1) = "1.1"
xlHoja1.Cells(7, 1) = "1.2"
xlHoja1.Cells(8, 1) = "1.3"
xlHoja1.Cells(9, 1) = "1.4"
xlHoja1.Cells(10, 1) = "1.5"
xlHoja1.Cells(11, 1) = "1.6"
xlHoja1.Cells(12, 1) = "1.7"
xlHoja1.Cells(13, 1) = "1.8"
xlHoja1.Cells(14, 1) = "1.9"

xlHoja1.Cells(6, 2) = "MONTO DEL SALDO DE COLOCACIONES DEL MES ANTERIOR"
xlHoja1.Cells(7, 2) = "N° DE COLOCACIONES DEL MES ANTERIOR"
xlHoja1.Cells(8, 2) = "N° DE NUEVAS COLOCACIONES"
xlHoja1.Cells(9, 2) = "MONTO DE NUEVAS COLOCACIONES"
xlHoja1.Cells(10, 2) = "MONTO DE AMORTIZACIONES"
xlHoja1.Cells(11, 2) = "N° DE CREDITOS AMORTIZADOS"
xlHoja1.Cells(12, 2) = "MONTO DEL SALDO DE COLOCACIONES (1.1+1.4-1.5)"
xlHoja1.Cells(13, 2) = "N° DE CREDITOS, SALDO DE COLOCACIONES (1.2+1.3+1.6)"
xlHoja1.Cells(14, 2) = "MONTO DE CUAOTAS QUE VENCEN EN EL MES SIGUIENTE"


xlHoja1.Range("A15").Cells.Font.Bold = True
xlHoja1.Range("A15").Cells.Font.Size = 10
xlHoja1.Cells(15, 1) = "2.   CARTERA ACTIVA"

xlHoja1.Cells(16, 1) = "2.1"
xlHoja1.Cells(17, 1) = "2.2"
xlHoja1.Cells(18, 1) = "2.3"
xlHoja1.Cells(19, 1) = "2.4"
xlHoja1.Cells(20, 1) = "2.5"
xlHoja1.Cells(21, 1) = "2.6"
xlHoja1.Cells(22, 1) = "2.7"
xlHoja1.Cells(23, 1) = "2.8"
xlHoja1.Cells(24, 1) = "2.9"
xlHoja1.Cells(25, 1) = "2.10"

xlHoja1.Cells(16, 2) = "MONTO DEL SALDO DEL MES ANTERIOR"
xlHoja1.Cells(17, 2) = "N° DE COLOCACIONES DEL MES ANTERIOR"
xlHoja1.Cells(18, 2) = "MONTO DE NUEVAS COLOCACIONES"
xlHoja1.Cells(19, 2) = "N° DE NUEVAS COLOCACIONES"
xlHoja1.Cells(20, 2) = "MONTO DE AMORTIZACIONES"
xlHoja1.Cells(21, 2) = "N° DE CREDITOS AMORTIZADOS"
xlHoja1.Range("C21:H21").MergeCells = True
xlHoja1.Cells(22, 2) = "MONTO DE PRESTAMOS CANCELADOS"
xlHoja1.Cells(23, 2) = "N° DE PRESTAMOS CANCELADOS"
xlHoja1.Cells(24, 2) = "MONTO DE SALDO DE COLOCACIONES [(2.1+2.3)-(2.5+2.7)]"
xlHoja1.Cells(25, 2) = "SALDO DEL N° DE COLOCACIONES (2.2+2.4-(2.6+2.8))"

xlHoja1.Range("A26").Cells.Font.Bold = True
xlHoja1.Range("A26").Cells.Font.Size = 10
xlHoja1.Cells(26, 1) = "3.   CARTERA MOROSA"

xlHoja1.Cells(27, 1) = "3.1"
xlHoja1.Cells(28, 1) = "3.2"
xlHoja1.Cells(29, 1) = "3.3"
xlHoja1.Cells(30, 1) = "3.4"
xlHoja1.Cells(31, 1) = "3.5"

xlHoja1.Cells(27, 2) = "CARTERA MOROSA DEL MES ANTERIOR"
xlHoja1.Cells(28, 2) = "MONTO DE CARTERA MOROSA"
xlHoja1.Cells(29, 2) = "N° DE CREDITOS MOROSOS"
xlHoja1.Cells(30, 2) = "CARTERA CONTAMINADA"
xlHoja1.Range("C30:H30").MergeCells = True
xlHoja1.Cells(31, 2) = "INDICE DE MORA (3.1/1.7)"

xlHoja1.Range("A32").Cells.Font.Bold = True
xlHoja1.Range("A32").Cells.Font.Size = 10
xlHoja1.Cells(32, 1) = "4.   CARTERA REFINANCIADA"

xlHoja1.Cells(33, 1) = "4.1"
xlHoja1.Cells(34, 1) = "4.2"
xlHoja1.Cells(35, 1) = "4.3"
xlHoja1.Cells(36, 1) = "4.4"
xlHoja1.Cells(37, 1) = "4.5"
xlHoja1.Cells(38, 1) = "4.6"
xlHoja1.Cells(39, 1) = "4.7"
xlHoja1.Cells(40, 1) = "4.8"
xlHoja1.Cells(41, 1) = "4.9"


xlHoja1.Cells(33, 2) = "MONTO DEL SALDO DEL MES ANTERIOR"
xlHoja1.Cells(34, 2) = "N° DE CREDITOS REFINANCIADOS"
xlHoja1.Cells(35, 2) = "MONTO DE CREDITOS REFINANCIADOS"
xlHoja1.Cells(36, 2) = "INTERES REFINANCIADO"
xlHoja1.Cells(37, 2) = "AMORTIZACIONES DE CREDITOS REFINANCIADOS"
xlHoja1.Cells(38, 2) = "N° DE CREDITOS REFINANCIADOS CANCELADOS"
xlHoja1.Cells(39, 2) = "MONTO DE CREDITOS REFINANCIADOS CANCELADOS"
xlHoja1.Cells(40, 2) = "SALDO DE MONTO DE CREDITOS REFIANCIADOS"
xlHoja1.Cells(41, 2) = "SALDO N° DE CREDITOS REFINANCIADOS"


xlHoja1.Range("A1").ColumnWidth = 5
xlHoja1.Range("B1").ColumnWidth = 55
xlHoja1.Range("B6:B41").Cells.Font.Size = 9

xlHoja1.Range("C6:H14").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("C6:H14").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("C6:H14").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("C6:H14").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("C6:H14").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("C6:H14").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("C6:H14").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("C6:H14").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

xlHoja1.Range("C16:H25").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("C16:H25").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("C16:H25").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("C16:H25").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("C16:H25").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("C16:H25").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("C16:H25").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("C16:H25").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

xlHoja1.Range("C27:H31").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("C27:H31").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("C27:H31").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("C27:H31").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("C27:H31").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("C27:H31").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("C27:H31").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("C27:H31").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

xlHoja1.Range("C33:H41").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("C33:H41").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("C33:H41").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("C33:H41").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("C33:H41").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("C33:H41").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("C33:H41").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("C33:H41").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

xlHoja1.Range("C6:H7").Cells.Interior.ColorIndex = 15
xlHoja1.Range("C6:H7").Cells.Interior.Pattern = xlSolid
xlHoja1.Range("C13:H14").Cells.Interior.ColorIndex = 15
xlHoja1.Range("C13:H14").Cells.Interior.Pattern = xlSolid
xlHoja1.Range("C16:H17").Cells.Interior.ColorIndex = 15
xlHoja1.Range("C16:H17").Cells.Interior.Pattern = xlSolid
xlHoja1.Range("C24:H25").Cells.Interior.ColorIndex = 15
xlHoja1.Range("C24:H25").Cells.Interior.Pattern = xlSolid
xlHoja1.Range("C33:H33").Cells.Interior.ColorIndex = 15
xlHoja1.Range("C33:H33").Cells.Interior.Pattern = xlSolid
xlHoja1.Range("C40:H41").Cells.Interior.ColorIndex = 15
xlHoja1.Range("C40:H41").Cells.Interior.Pattern = xlSolid
' ---------------->  Cartera Total


' ---------------->  Cartera Activa

J = 3
' Nuevos Cred y Cap Amortizado
For i = 0 To 5
    Set rs = oCapRCD.ObtenerNuevoCredCapArmotizado(sFechaIni)
    If rs.EOF And rs.BOF Then
    Else
        xlHoja1.Cells(J, 18) = rs!MontoDesembN
        xlHoja1.Cells(J, 19) = rs!NroDesembN
        xlHoja1.Cells(J, 22) = rs!CapPag
        If J + 1 < 8 Then xlHoja1.Cells(J + 1, 27) = rs!MontoDesembN
    
    End If
    J = J + 1
Next i
rs.Close

' Monto de Prestamos Cancelados
J = 3
Set rs = New ADODB.Recordset
For i = 0 To 5
    Set rs = oCapRCD.ObtenerMontoPrestamos(sFechaIni)
    If rs.EOF And rs.BOF Then
    Else
        xlHoja1.Cells(J, 22) = rs!MontoDesembN
        xlHoja1.Cells(J, 23) = rs!NroDesembN
    End If
    J = J + 1
Next i
rs.Close

For i = Asc("C") To Asc("H")
    xlHoja1.Range(Chr(i) & "24").Formula = "=(" & Chr(i) & "16 +" & Chr(i) & "18 ) - (" & Chr(i) & "20 +" & Chr(i) & "22 )"
    xlHoja1.Range(Chr(i) & "25").Formula = "=(" & Chr(i) & "17 +" & Chr(i) & "19 ) - (" & Chr(i) & "23)"
    If Chr(i) <> "G" Then
        xlHoja1.Cells(Chr(i + 1), 16) = xlHoja1.Cells(Chr(i) & "24")
        xlHoja1.Cells(Chr(i + 1), 17) = xlHoja1.Cells(Chr(i) & "25")
    End If
Next i

' ---------------->  Cartera Morosa
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerCarteraMoroso(sFechaIni, -1)
xlHoja1.Cells(3, 27) = rs!MontoJud
rs.Close
J = 3
Set rs = New ADODB.Recordset
For i = 0 To 5
    Set rs = oCapRCD.ObtenerCarteraMoroso(sFechaIni, i)
    Set rs = GetQuery(sql)
    If rs.EOF And rs.BOF Then
    Else
        xlHoja1.Cells(J, 28) = rs!MontoJud
        xlHoja1.Cells(J, 29) = rs!NroJudi
        If J + 1 < 8 Then xlHoja1.Cells(J + 1, 27) = rs!MontoJud
    End If
    J = J + 1
Next i
rs.Close

For i = Asc("C") To Asc("H")
    xlHoja1.Range(Chr(i) & "24").Formula = "=(" & Chr(i) & "16 +" & Chr(i) & "18 ) - (" & Chr(i) & "20 +" & Chr(i) & "22 )"
    xlHoja1.Range(Chr(i) & "25").Formula = "=(" & Chr(i) & "17 +" & Chr(i) & "19 ) - (" & Chr(i) & "23)"
    If Chr(i) <> "G" Then
        xlHoja1.Cells(Chr(i + 1), 40) = "= " & Chr(i) & "33 + " & Chr(i) & "35 +" & Chr(i) & "37"
        xlHoja1.Cells(Chr(i + 1), 41) = "= " & Chr(i) & "34 + " & Chr(i) & "38"
    End If
Next i

' ---------------->  Cartera Refinanciada
J = 3
Set rs = New ADODB.Recordset
For i = 0 To 5
    Set rs = oCapRCD.ObtenerCarteraRefinanciada(sFechaIni)
    If rs.EOF And rs.BOF Then
    Else
        xlHoja1.Cells(J, 34) = rs!NroRef
        xlHoja1.Cells(J, 35) = rs!MonoRefinan
        xlHoja1.Cells(J, 36) = rs!nIntPag
        xlHoja1.Cells(J, 37) = rs!nCapPag
        xlHoja1.Cells(J, 38) = rs!NroCan
        xlHoja1.Cells(J, 39) = rs!SaldoCap
        If J + 1 < 8 Then xlHoja1.Cells(J + 1, 33) = rs!MontoJud
    End If
    J = J + 1
Next i
rs.Close
Set oCapRCD = Nothing
xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'RaiseEvent Progress(20, 20)
'---------------->
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
'-------------->
'RaiseEvent CloseProgress
MsgBox "Se ha Generado el Archivo " & lsArchivo & ".XLS Satisfactoriamente", vbInformation, "Aviso"
Exit Sub
ErrorExcel:
    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
    xlLibro.Close
    ' Cierra Microsoft Excel con el método Quit.
    xlAplicacion.Quit
    'Libera los objetos.
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing
End Sub

Public Sub nRepo_IBM_Excel(ByVal psServerConsol As String, ByVal pdFechaCierre As Date, ByVal psNomCmac As String, dData As Date)
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim fs As Scripting.FileSystemObject

Dim sql As String
Dim rs As New ADODB.Recordset

Dim lsArchivo As String
Dim nFil As Integer
Dim nCol As Integer
Dim lsCodigo As String
Dim Titulo As String
Dim i As Integer
Dim J As Integer
Dim oConst As COMDConstSistema.DCOMConstSistema
Dim oCapRCD As COMDCaptaGenerales.DCOMRCD

Set oConst = New COMDConstSistema.DCOMConstSistema
    Set rs = oConst.ObtenerVarSistema80()
    lsCodigo = IIf(IsNull(rs!nConsSisValor), "", rs!nConsSisValor)
Set oConst = Nothing

lsArchivo = "ResumenIBM" & Format(pdFechaCierre, "YYYYMMDD") & ".xls"
Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

Set xlHoja1 = xlLibro.Worksheets.Add
'RaiseEvent Progress(2, 20)

xlHoja1.Range("A1").Cells.Font.Bold = True
xlHoja1.Range("A1").Cells.Font.Size = 10
xlHoja1.Cells(1, 1) = "SUPERINTENDENCIA DE BANCA Y SEGUROS" 'psNomCmac
xlHoja1.Range("A2").Cells.Font.Size = 8
xlHoja1.Cells(2, 1) = "Departamento de Evaluacion de Riesgos Crediticio"

xlHoja1.Range("F1:G2").Cells.Font.Size = 8
xlHoja1.Cells(1, 6) = "Periodo"
xlHoja1.Cells(2, 6) = "Fecha"
xlHoja1.Cells(1, 7) = Format(dData, "mmm yy")
xlHoja1.Range("F1:G2").HorizontalAlignment = xlRight
xlHoja1.Cells(2, 7) = dData

xlHoja1.Range("A6:G6").MergeCells = True
xlHoja1.Cells(6, 1) = "RESUMEN DE TOTALES POR ITEM"
xlHoja1.Range("A6").HorizontalAlignment = xlCenter
xlHoja1.Range("A6").Cells.Font.Bold = True
xlHoja1.Range("A6").Cells.Font.Size = 15

xlHoja1.Cells(8, 2) = "Empresa"
xlHoja1.Cells(9, 2) = "Código"
xlHoja1.Cells(10, 2) = "N° Deudores Reportados"

xlHoja1.Range("C8:F8").MergeCells = True
xlHoja1.Range("C8:F8").Cells.BorderAround xlContinuous
xlHoja1.Cells(8, 3) = psNomCmac
xlHoja1.Range("C9:F9").MergeCells = True
xlHoja1.Range("C9:F9").Cells.BorderAround xlContinuous
xlHoja1.Range("C9").HorizontalAlignment = xlCenter

xlHoja1.Cells(9, 3) = lsCodigo
'xlHoja1.Cells(10, 3) = "N° Deudores Reportados"

xlHoja1.Range("A13:G13").Cells.Font.Bold = True
xlHoja1.Range("A13:G13").Cells.Font.Size = 8
xlHoja1.Range("A13:G13").HorizontalAlignment = xlCenter
xlHoja1.Cells(13, 1) = "Item"
xlHoja1.Cells(13, 3) = "Cantidad"
xlHoja1.Cells(13, 5) = "Item"
xlHoja1.Cells(13, 7) = "Cantidad"

J = 15
For i = 3 To 38
    xlHoja1.Cells(J, 1) = i
    J = J + 1
Next i


J = 15
Set oCapRCD = New COMDCaptaGenerales.DCOMRCD
Dim sValor As String

Set rs = New ADODB.Recordset
For i = 1 To 6
    Set rs = oCapRCD.ObtenerIBM1(psServerConsol, i)
    xlHoja1.Cells(J, 3) = rs!Total
    J = J + 1
Next i
rs.Close

xlHoja1.Cells(15, 2) = "NUM DEU CON DOCUM T/1"
xlHoja1.Cells(16, 2) = "NUM DEU CON DOCUM T/2"
xlHoja1.Cells(17, 2) = "NUM DEU CON DOCUM T/3"
xlHoja1.Cells(18, 2) = "NUM DEU CON DOCUM T/4"
xlHoja1.Cells(19, 2) = "NUM DEU CON DOCUM T/5"
xlHoja1.Cells(20, 2) = "NUM DEU CON DOCUM T/6"
'RaiseEvent Progress(2, 20)
J = 21
Set rs = New ADODB.Recordset
For i = 1 To 2
    Set rs = oCapRCD.ObtenerIBM1(psServerConsol, i)
    xlHoja1.Cells(J, 3) = rs!Total
    J = J + 1
Next i
rs.Close
xlHoja1.Cells(21, 2) = "NUM DEU PERSONA NATURAL"
xlHoja1.Cells(22, 2) = "NUM DEU PERSONA JURIDICA"

'**************
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenernVen30MN(psServerConsol)
xlHoja1.Cells(23, 3) = rs!Total
xlHoja1.Cells(23, 2) = "NUM DEU CUO VENC < 30 MN"
rs.Close

Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenernVen31MN(psServerConsol)
xlHoja1.Cells(24, 3) = rs!Total
xlHoja1.Cells(24, 2) = "NUM DEU CUO VENC > 30 MN"
rs.Close

Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenernVen30MN(psServerConsol)
xlHoja1.Cells(25, 3) = rs!Total
xlHoja1.Cells(25, 2) = "NUM DEU CUO VENC < 30 ME"
rs.Close

Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenernVen31MN(psServerConsol)
xlHoja1.Cells(26, 3) = rs!Total
xlHoja1.Cells(26, 2) = "NUM DEU CUO VENC > 30 ME"
'****************
rs.Close

Set rs = New ADODB.Recordset
J = 27
For i = 0 To 4
    Set rs = oCapRCD.ObtenerTotalCalifica(psServer, i)
    xlHoja1.Cells(J, 3) = rs!Total
    J = J + 1
Next i
rs.Close

xlHoja1.Cells(27, 2) = "NUM DEU CON CALIF 0"
xlHoja1.Cells(28, 2) = "NUM DEU CON CALIF 1"
xlHoja1.Cells(29, 2) = "NUM DEU CON CALIF 2"
xlHoja1.Cells(30, 2) = "NUM DEU CON CALIF 3"
xlHoja1.Cells(31, 2) = "NUM DEU CON CALIF 4"
'*******************
'RaiseEvent Progress(6, 20)
xlHoja1.Cells(32, 2) = "NUM DEU CON P/EFECT AVAL"
xlHoja1.Cells(32, 3) = "0"
xlHoja1.Cells(33, 2) = "NUM DEU CON P/EFECT TERCERO"
xlHoja1.Cells(33, 3) = "0"

'******************
'nRefME,nVen30ME,nVen31ME,nCobJudME,nLinCredME,nCastME

sValor = "nNormMN"
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerDeudaVigente(psServer, sValor)
xlHoja1.Cells(34, 3) = "=" & rs!Total
xlHoja1.Cells(34, 2) = "MN DEUDA DIRECTA VIGENTE"
rs.Close

xlHoja1.Cells(35, 3) = "0"
xlHoja1.Cells(35, 2) = "MN DD ARREND FINACIERO"

sValor = "nRefMN"
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerDeudaDirecRefinanciada(psServer, sValor)
xlHoja1.Cells(36, 3) = "=" & rs!Total
xlHoja1.Cells(36, 2) = "MN DEUDA DIREC REFINANCIADA"
rs.Close

sValor = "nVen30MN"
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerDeudaDirecVencida30(psServer, sValor)
xlHoja1.Cells(37, 3) = "=" & rs!Total
xlHoja1.Cells(37, 2) = "MN DEUDA DIREC VENC < 30"
rs.Close

sValor = "nVen31MN"
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerDeudaDirecVencida31(psServer, sValor)
xlHoja1.Cells(38, 3) = "=" & rs!Total
xlHoja1.Cells(38, 2) = "MN DEUDA DIREC VENC > 30"
rs.Close

sValor = "nCobJudMN"
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerDeudaDirecJudicial(psServer, sValor)
xlHoja1.Cells(39, 3) = "=" & rs!Total
xlHoja1.Cells(39, 2) = "MN DEUDA DIREC COB JUDICIAL"

xlHoja1.Cells(40, 2) = "MN DEUDA INDIRECTA"
xlHoja1.Cells(40, 3) = "0"
xlHoja1.Cells(41, 2) = "MN DEUDA AVALADA"
xlHoja1.Cells(41, 3) = "0"
rs.Close

sValor = "nLinCredMN"
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerLineasCredito(psServer, sValor)
xlHoja1.Cells(42, 3) = "=" & rs!Total
xlHoja1.Cells(42, 2) = "MN LINEA DE CREDITO"
rs.Close

sValor = "nCastMN"
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerCreditosCastigados(psServer, sValor)
xlHoja1.Cells(43, 3) = "=" & rs!Total
xlHoja1.Cells(43, 2) = "MN CREDITOS CASTIGADOS"
rs.Close

xlHoja1.Cells(44, 2) = "MN DEUDA VENDIDA A PLAZO"
xlHoja1.Cells(44, 3) = "0"
xlHoja1.Cells(45, 2) = "MN DEUDA VENDIDA AL CONTADO"
xlHoja1.Cells(45, 3) = "0"
xlHoja1.Cells(46, 2) = "MN DEUDA B / ADM VIGENTE"
xlHoja1.Cells(46, 3) = "0"
xlHoja1.Cells(47, 2) = "MN DEUDA B / ADM REFINANCIADA"
xlHoja1.Cells(47, 3) = "0"
xlHoja1.Cells(48, 2) = "MN DEUDA B / ADM VENC < 30"
xlHoja1.Cells(48, 3) = "0"
xlHoja1.Cells(49, 2) = "MN DEUDA B / ADM VEMC > 30"
xlHoja1.Cells(49, 3) = "0"
xlHoja1.Cells(50, 2) = "MN DEUDA B / ADM COB JUDICIAL"
xlHoja1.Cells(50, 3) = "0"

' Dolares
J = 39
For i = 34 To 50
    xlHoja1.Cells(i, 5) = J
    J = J + 1
Next i

'*************************
'RaiseEvent Progress(9, 20)

sValor = "nNormME"
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerDeudaVigente(psServer, sValor)
xlHoja1.Cells(34, 7) = rs!Total
xlHoja1.Cells(34, 6) = "ME DEUDA DIRECTA VIGENTE"
rs.Close

xlHoja1.Cells(35, 7) = "0"
xlHoja1.Cells(35, 6) = "ME DD ARREND FINANCIERO"

sValor = "nRefME"
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerDeudaDirecRefinanciada(psServer, sValor)
xlHoja1.Cells(36, 7) = rs!Total
xlHoja1.Cells(36, 6) = "ME DEUDA DIREC REFINANCIADA"
rs.Close

sValor = "nVen30ME"
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerDeudaDirecVencida30(psServer, sValor)
xlHoja1.Cells(37, 7) = rs!Total
xlHoja1.Cells(37, 6) = "ME DEUDA DIREC VENC < 30"
rs.Close

sValor = "nVen31ME"
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerDeudaDirecVencida31(psServer, sValor)
xlHoja1.Cells(38, 7) = rs!Total
xlHoja1.Cells(38, 6) = "ME DEUDA DIREC VENC > 30"
rs.Close

sValor = "nCobJudME"
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerDeudaDirecJudicial(psServer, sValor)
xlHoja1.Cells(39, 7) = rs!Total
xlHoja1.Cells(39, 6) = "ME DEUDA DIREC COB JUDICIAL"
rs.Close

xlHoja1.Cells(40, 7) = "0"
xlHoja1.Cells(40, 6) = "ME DEUDA INDIRECTA"
'RaiseEvent Progress(11, 20)
xlHoja1.Cells(41, 7) = "0"
xlHoja1.Cells(41, 6) = "ME DEUDA LAVADA"

sValor = "nLinCredME"
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerLineasCredito(psServer, sValor)
xlHoja1.Cells(42, 7) = rs!Total
xlHoja1.Cells(42, 6) = "ME DEUDA LINEA DE CREDITO"
rs.Close

sValor = "nCastME"
Set rs = New ADODB.Recordset
Set rs = oCapRCD.ObtenerCreditosCastigados(psServer, sValor)
xlHoja1.Cells(43, 7) = rs!Total
xlHoja1.Cells(43, 6) = "ME DEUDA CREDITOS CASTIGADOS"
rs.Close

xlHoja1.Cells(44, 7) = "0"
xlHoja1.Cells(44, 6) = "ME DEUDA VENDIDA A PLAZO"
'RaiseEvent Progress(12, 20)
xlHoja1.Cells(45, 7) = "0"
xlHoja1.Cells(45, 6) = "ME DEUDA VENDIDA AL CONTADO"
'************************
xlHoja1.Cells(46, 6) = "ME DEUDA B / ADM VIGENTE"
xlHoja1.Cells(46, 7) = "0"
xlHoja1.Cells(47, 6) = "ME DEUDA B / ADM REFINANCIADA"
xlHoja1.Cells(47, 7) = "0"
xlHoja1.Cells(48, 6) = "ME DEUDA B / ADM VENC < 30"
xlHoja1.Cells(48, 7) = "0"
xlHoja1.Cells(49, 6) = "ME DEUDA B / ADM VEMC > 30"
xlHoja1.Cells(49, 7) = "0"
xlHoja1.Cells(50, 6) = "ME DEUDA B / ADM COB JUDICIAL"
xlHoja1.Cells(50, 7) = "0"
xlHoja1.Range("A15:G70").Cells.Font.Size = 8

xlHoja1.Range("A1").ColumnWidth = 4
xlHoja1.Range("B1").ColumnWidth = 25
xlHoja1.Range("C1").ColumnWidth = 17
xlHoja1.Range("D1").ColumnWidth = 3
xlHoja1.Range("E1").ColumnWidth = 4
xlHoja1.Range("F1").ColumnWidth = 25
xlHoja1.Range("G1").ColumnWidth = 17

xlHoja1.Cells(55, 2) = "(03+04+05+06+07+08)=(09+10)"
xlHoja1.Cells(56, 2) = "'=(15+16+17+18+19)"
xlHoja1.Cells(57, 2) = "(22+23+24)"
xlHoja1.Cells(58, 2) = "(25+26+27)"
xlHoja1.Cells(59, 2) = "(39+40+41)"
xlHoja1.Cells(60, 2) = "(42+43+44)"

xlHoja1.Cells(62, 2) = "(a+c)"
xlHoja1.Cells(63, 2) = "(b+d)"
xlHoja1.Cells(64, 2) = "(a+b+c+d)"

xlHoja1.Cells(66, 2) = "(28+45)"
xlHoja1.Cells(67, 2) = "(A+B)"
'--------------->
'RaiseEvent Progress(14, 20)
xlHoja1.Range("C15:C20").Cells.BorderAround xlContinuous
xlHoja1.Range("C23:C26").Cells.BorderAround xlContinuous
xlHoja1.Range("C21:C22").Cells.BorderAround xlContinuous
xlHoja1.Range("C27:C31").Cells.BorderAround xlContinuous
xlHoja1.Range("C32:C33").Cells.BorderAround xlContinuous
xlHoja1.Range("C34:C39").Cells.BorderAround xlContinuous
xlHoja1.Range("C40:C45").Cells.BorderAround xlContinuous
xlHoja1.Range("C46:C50").Cells.BorderAround xlContinuous
xlHoja1.Range("G34:G39").Cells.BorderAround xlContinuous
xlHoja1.Range("G40:G45").Cells.BorderAround xlContinuous
xlHoja1.Range("G46:G50").Cells.BorderAround xlContinuous
xlHoja1.Range("F55").Cells.BorderAround xlContinuous
xlHoja1.Range("F57:F60").Cells.BorderAround xlContinuous
xlHoja1.Range("F62:F64").Cells.BorderAround xlContinuous
xlHoja1.Range("F66").Cells.BorderAround xlContinuous
xlHoja1.Range("F67").Cells.BorderAround xlContinuous, xlMedium
xlHoja1.Range("F67").Font.Bold = True

xlHoja1.Cells(55, 3) = "NUM DE DEUDORES"
xlHoja1.Range("F55").Formula = "=SUM(C21..C22)"  ' NUM DEUDORES

xlHoja1.Cells(57, 3) = "MN DEUDA VIGENTE (a)"
xlHoja1.Range("F57").Formula = "=SUM(C34..C36)"  ' (a)

xlHoja1.Cells(58, 3) = "MN DEUDA VENCIDA (b)"
xlHoja1.Range("F58").Formula = "=SUM(C37..C39)"  ' (b)

xlHoja1.Cells(59, 3) = "ME DEUDA VIGENTE (c)"
xlHoja1.Range("F59").Formula = "=SUM(G34..G36)"  ' (c)

xlHoja1.Cells(60, 3) = "ME DEUDA VENCIDA (d)"
xlHoja1.Range("F60").Formula = "=SUM(G37..G39)"  ' (d)
'-------------->
'RaiseEvent Progress(17, 20)
xlHoja1.Cells(62, 3) = "DEUDA VIGENTE"
xlHoja1.Range("F62").Formula = "=SUM(F57..F58)"

xlHoja1.Cells(63, 3) = "DEUDA VENCIDA"
xlHoja1.Range("F63").Formula = "=SUM(F59..F60)"

xlHoja1.Cells(64, 3) = "DEUDA DIRECTA (A)"
xlHoja1.Range("F64").Formula = "=SUM(F57..F60)"

xlHoja1.Cells(66, 3) = "DEUDA INDIRECTA (B)"
xlHoja1.Range("F66").Formula = "=C40+G45"

xlHoja1.Cells(67, 3) = "DEUDA TOTAL"
xlHoja1.Range("F67").Formula = "=F64+F66"
'RaiseEvent Progress(18, 20)

    
xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'RaiseEvent Progress(20, 20)
'---------------->
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
'Set oTipCambio = Nothing
'-------------->

'RaiseEvent CloseProgress

MsgBox "Se ha Generado el Archivo " & lsArchivo & ".XLS Satisfactoriamente", vbInformation, "Aviso"
Exit Sub
ErrorExcel:
    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
    xlLibro.Close
    ' Cierra Microsoft Excel con el método Quit.
    xlAplicacion.Quit
    'Libera los objetos.
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing
    
End Sub

