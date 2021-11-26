VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nRcdReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim Conecta As DConecta
Public Event ShowProgress()
Public Event CloseProgress()
Public Event Progress(pnValor As Long, pnTotal As Long)
Function GetQuery(ByVal sql As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Conecta.AbreConexion
Set rs = Conecta.CargaRecordSet(sql)
Conecta.CierraConexion
Set GetQuery = rs
End Function
Public Function GetServerConsol() As String
Dim rs As New ADODB.Recordset
Set rs = GetQuery("select nConsSisValor from ConstSistema where nConsSisCod=43")
GetServerConsol = IIf(IsNull(rs!nConsSisValor), "", rs!nConsSisValor)
Set rs = Nothing
End Function
Public Function ExisteTabla(ByVal psServerConsol As String, ByVal psTabla As String) As ADODB.Recordset
Dim sql As String
Dim rs As New ADODB.Recordset
sql = "select * from " & psServerConsol & "sysobjects where name = '" & psTabla & "'"
Set rs = GetQuery(sql)
Set ExisteTabla = rs
Set rs = Nothing
End Function
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
Dim I As Integer

Dim J As Integer

Dim Mes(1 To 12) As String
Dim nMes As Integer
Dim sFechaIni As String
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
For I = 0 To 5
    sql = " select isnull(sum(nNumDesembN),0) NroDesembN ,"
    sql = sql & "       isnull(sum(nMontoDesembN) ,0) MontoDesembN,"
    sql = sql & "       isnull(sum(nCapPag),0) CapPag"
    sql = sql & " From ColocEstadDiaCred"
    sql = sql & " where datediff(m,dEstad,'" & Format(DateAdd("m", I, sFechaIni), "yyyy/mm/dd") & "')=0"
    sql = sql & " and cLineaCred='02'"
    Set rs = GetQuery(sql)
    If rs.EOF And rs.BOF Then
    Else
        xlHoja1.Cells(J, 18) = rs!MontoDesembN
        xlHoja1.Cells(J, 19) = rs!NroDesembN
        xlHoja1.Cells(J, 22) = rs!CapPag
        If J + 1 < 8 Then xlHoja1.Cells(J + 1, 27) = rs!MontoDesembN
    
    End If
    J = J + 1
Next I
' Monto de Prestamos Cancelados
J = 3
For I = 0 To 5
    sql = "       select isnull(sum(nNumCanc),0) NroCanc ,"
    sql = sql & "        isnull(sum(nSaldoCap) ,0) SaldoCap"
    sql = sql & " From ColocEstadDiaCred"
    sql = sql & " where datediff(m,dEstad,'" & Format(DateAdd("m", I, sFechaIni), "yyyy/mm/dd") & "')=0"
    sql = sql & " and cLineaCred='02'"
    Set rs = GetQuery(sql)
    If rs.EOF And rs.BOF Then
    Else
        xlHoja1.Cells(J, 22) = rs!MontoDesembN
        xlHoja1.Cells(J, 23) = rs!NroDesembN
    End If
    J = J + 1
Next I


For I = Asc("C") To Asc("H")
    xlHoja1.Range(Chr(I) & "24").Formula = "=(" & Chr(I) & "16 +" & Chr(I) & "18 ) - (" & Chr(I) & "20 +" & Chr(I) & "22 )"
    xlHoja1.Range(Chr(I) & "25").Formula = "=(" & Chr(I) & "17 +" & Chr(I) & "19 ) - (" & Chr(I) & "23)"
    If Chr(I) <> "G" Then
        xlHoja1.Cells(Chr(I + 1), 16) = xlHoja1.Cells(Chr(I) & "24")
        xlHoja1.Cells(Chr(I + 1), 17) = xlHoja1.Cells(Chr(I) & "25")
    End If
Next I

' ---------------->  Cartera Morosa
sql = " select isnull(sum(nNumJud),0) NroJudi , isnull(sum(nMontoJudic) ,0) MontoJud "
sql = sql & " From ColocEstadDiaCred "
sql = sql & " where datediff(m,dEstad,'" & Format(DateAdd("m", -1, sFechaIni), "yyyy/mm/dd") & "')=0 "
sql = sql & " and cLineaCred='02'"
Set rs = GetQuery(sql)
xlHoja1.Cells(3, 27) = rs!MontoJud
J = 3
For I = 0 To 5
    sql = " select isnull(sum(nNumJud),0) NroJudi , isnull(sum(nMontoJudic) ,0) MontoJud "
    sql = sql & " From ColocEstadDiaCred "
    sql = sql & " where datediff(m,dEstad,'" & Format(DateAdd("m", I, sFechaIni), "yyyy/mm/dd") & "')=0 "
    sql = sql & " and cLineaCred='02'"
    Set rs = GetQuery(sql)
    If rs.EOF And rs.BOF Then
    Else
        xlHoja1.Cells(J, 28) = rs!MontoJud
        xlHoja1.Cells(J, 29) = rs!NroJudi
        If J + 1 < 8 Then xlHoja1.Cells(J + 1, 27) = rs!MontoJud
    End If
    J = J + 1
Next I

For I = Asc("C") To Asc("H")
    xlHoja1.Range(Chr(I) & "24").Formula = "=(" & Chr(I) & "16 +" & Chr(I) & "18 ) - (" & Chr(I) & "20 +" & Chr(I) & "22 )"
    xlHoja1.Range(Chr(I) & "25").Formula = "=(" & Chr(I) & "17 +" & Chr(I) & "19 ) - (" & Chr(I) & "23)"
    If Chr(I) <> "G" Then
        xlHoja1.Cells(Chr(I + 1), 40) = "= " & Chr(I) & "33 + " & Chr(I) & "35 +" & Chr(I) & "37"
        xlHoja1.Cells(Chr(I + 1), 41) = "= " & Chr(I) & "34 + " & Chr(I) & "38"
    End If
Next I

' ---------------->  Cartera Refinanciada
J = 3
For I = 0 To 5
    sql = "       Select isnull(sum(nNumRef),0) NroRef ,"
    sql = sql & "       isnull(sum(nMontoRefinan) ,0) MonoRefinan,"
    sql = sql & "       isnull(sum(nIntPag),0) nIntPag,"
    sql = sql & "       isnull(sum(nCapPag),0) nCapPag,"
    sql = sql & "       isnull(sum(nNumCanc),0) NroCan,"
    sql = sql & "       isnull(sum(nSaldoCap),0) SaldoCap"
    sql = sql & " From ColocEstadDiaCred"
    sql = sql & " where datediff(m,dEstad,'" & Format(DateAdd("m", I, sFechaIni), "yyyy/mm/dd") & "')=0"
    sql = sql & " and cLineaCred='02'"
    Set rs = GetQuery(sql)
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
Next I

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
Dim I As Integer
Dim J As Integer

Dim Mes(1 To 12) As String
Dim nMes As Integer
Dim sFechaIni As String
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
sql = sql & "SELECT Fecha, Count(cCtaCod) Numero, Sum(Monto) Monto from ("
sql = sql & " Select Convert(Varchar(7),dAsignacion,102) Fecha, cCtaCod, Monto = CASE"
sql = sql & "    WHEN nPrdEstado In ('2000','2003') THEN nMontoSol"
sql = sql & "    WHEN nPrdEstado In ('2001') THEN nMontoSug"
sql = sql & "    WHEN nPrdEstado In ('2002','2201','2080') THEN nMontoApr"
sql = sql & "    WHEN nPrdEstado In ('2020','2021','2022','2030','2031','2032','2050') THEN nMontoDesemb END"
sql = sql & " From " & psServConsol & "CreditoConsolTotal"
sql = sql & " where SubString(cLineaCred,1,2) = '02'"
sql = sql & " And nPrdEstado In ('2000','2003','2001','2002','2201','2080','2020','2021','2022','2030','2031','2032','2050')"
sql = sql & " and DateDiff(dd,dAsignacion,'" & Format(sFechaIni, "mm/dd/yyyy") & "') <= 0  And DateDiff(dd,dAsignacion,'" & Format(pdFechaCierre, "mm/dd/yyyy") & "') >= 0"
sql = sql & " ) C Group by Fecha Order by Fecha"

Set rs = GetQuery(sql)
I = 8
While Not rs.EOF
    xlHoja1.Cells(9, I) = rs!Numero
    xlHoja1.Cells(10, I) = Format(rs!Monto, "0.00")
    I = I - 1
    rs.MoveNext
Wend


' CREDITOS APROBADOS
sql = "SELECT Fecha, Count(cCtaCod) Numero, Sum(Monto) Monto from ("
sql = sql & " Select Convert(Varchar(7),dFecApr,102) Fecha, cCtaCod, Monto = CASE"
sql = sql & "     WHEN nPrdEstado In ('2002') THEN nMontoApr"
sql = sql & "     WHEN nPrdEstado In ('2050','2020','2021','2022','2030','2031','2032') THEN nMontoDesemb END"
sql = sql & " From " & psServConsol & "CreditoConsolTotal"
sql = sql & " where SubString(cLineaCred,1,2) = '02'"
sql = sql & " And nPrdEstado In ('2020','2021','2022','2050','2030','2031','2032','2002')"
sql = sql & " and DateDiff(dd,dFecApr,'" & Format(sFechaIni, "mm/dd/yyyy") & "') <= 0  And DateDiff(dd,dFecApr,'" & Format(pdFechaCierre, "mm/dd/yyyy") & "') >= 0"
sql = sql & " ) C Group by Fecha Order by Fecha"
Set rs = GetQuery(sql)
I = 8

While Not rs.EOF
    xlHoja1.Cells(12, I) = rs!Numero
    xlHoja1.Cells(13, I) = Format(rs!Monto, "0.00")
    I = I - 1
    rs.MoveNext
Wend

'   CREDITOS DESEMBOLSADOS
sql = " Select Convert(Varchar(7),dFecVig,102) Fecha, Count(cCtaCod) Numero, Sum(nMontoDesemb) Monto"
sql = sql & " From " & psServConsol & "CreditoConsolTotal where SubString(cLineaCred,1,2) = '02' And nPrdEstado In ('2201','2020','2021','2022','2050','2031','2032','2002')"
sql = sql & " and DateDiff(dd,dFecVig,'" & Format(sFechaIni, "mm/dd/yyyy") & "') <= 0  And DateDiff(dd,dFecVig,'" & Format(pdFechaCierre, "mm/dd/yyyy") & "') >= 0"
sql = sql & " Group by Convert(Varchar(7),dFecVig,102) Order by Fecha"
Set rs = GetQuery(sql)
I = 8
While Not rs.EOF
    xlHoja1.Cells(12, I) = rs!Numero
    xlHoja1.Cells(13, I) = Format(rs!Monto, "0.00")
    I = I - 1
    rs.MoveNext
Wend

'  PLAZOS

For I = 0 To 5
    sql = " Select Rango, Count(Cuenta) Num, isnull(Sum(Desembolso),0) Desembolso, Sum(Saldo) Saldo FROM ("
    sql = sql & " Select C.cuenta, C.Desembolso, C.Saldo, C.Plazo , Rango = CASE"
    sql = sql & "     WHEN C.Plazo <= 180 THEN '01. Hasta 6 Meses'"
    sql = sql & "     WHEN C.Plazo > 180 And Plazo <= 360 THEN '02. De 6 a 12 Meses'"
    sql = sql & "     WHEN C.Plazo > 360 And Plazo <= 720  THEN '03. De 12 a 24 Meses'"
    sql = sql & "     WHEN C.Plazo > 720  THEN '04. Mas de 24 Meses' END FROM ("
    sql = sql & " Select cCtaCod cuenta, nMontoDesemb Desembolso, nsaldocap saldo, ncuotasapr*nplazoapr plazo"
    sql = sql & " FROM " & psServConsol & "CreditoConsolTotal WHERE nPrdEstado IN ('2020','2021','2022','2050','2030','2031','2032','2002') And SubString(cLineaCred,1,2) = '02'"
    sql = sql & "  and nTipcuota in ('10','11','20','21','30','31')"
    sql = sql & "  and DateDiff(mm,dFecVig,'" & Format(DateAdd("m", I, sFechaIni), "yyyy/mm/dd") & "') = 0"
    sql = sql & " Union"
    sql = sql & " Select cCtaCod cuenta, nMontoDesemb Desembolso, nsaldocap saldo, ncuotasapr*30 plazo"
    sql = sql & " FROM " & psServConsol & "CreditoConsolTotal WHERE nPrdEstado IN ('2020','2021','2022','2050','2030','2031','2032','2002') and SubString(cLineaCred,1,2) = '02'"
    sql = sql & "  and nTipcuota in ('40','41','50','51','60','61','70','71')"
    sql = sql & "  and DateDiff(mm,dFecVig,'" & Format(DateAdd("m", I, sFechaIni), "yyyy/mm/dd") & "') = 0"
    sql = sql & " ) C"
    sql = sql & " ) C1 Group by Rango Order by Rango"
    Set rs = GetQuery(sql)
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
Next I

'''''''''''''

'''''''''''''


'    POR SEXO
For I = 0 To 5
    sql = " SELECT cSexPers, count(cCtaCod) Cuenta , isnull(sum(T.Desembolso),0) Desembolso FROM ("
    sql = sql & "    Select"
    sql = sql & "    cSexPers = CASE WHEN cPersNatSexo = 'M' THEN 'H'"
    sql = sql & "              WHEN cPersNatSexo = 'F' THEN 'M' ELSE 'J'"
    sql = sql & "               END,"
    sql = sql & "    C.cCtaCod , nMontoDesemb Desembolso"
    sql = sql & "    From " & psServConsol & "CreditoConsolTotal c"
    sql = sql & "    Inner Join " & psServConsol & "ProductoPersonaConsol PE on C.cCtaCod = Pe.cCtaCod"
    sql = sql & "    Inner Join Persona P on P.cPersCod = Pe.cPersCod"
    sql = sql & "    left outer join PersonaNat PN on PN.cPersCod = P.cPersCod"
    sql = sql & "    Where "
    sql = sql & "    C.nPrdEstado in ('2020','2021','2022','2050','2030','2031','2032') and substring(cLineaCred,1,2) = '02'"
    sql = sql & "    and PE.nPrdPersRelac = 20"
    sql = sql & "   and DateDiff(mm,dfecvig,'" & Format(DateAdd("m", I, sFechaIni), "yyyy/mm/dd") & "') = 0"
    sql = sql & " ) T"
    sql = sql & " GROUP BY T.cSexPers ORDER BY T.cSexPers"
    Set rs = GetQuery(sql)
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
Next I


' COLOCACIONES POR NUMERO DE TRABAJADORES
For I = 0 To 5
    sql = " Select Fecha, NroTraba, Count(Cuenta) Num, isnull(Sum(Desembolso),0) Desembolso, Sum(Saldo) Saldo FROM"
    sql = sql & " (  Select Fecha, C.cuenta, C.Desembolso, C.Saldo, NroTRaba = CASE"
    sql = sql & "   WHEN nNroEmp <= 10 THEN '01. Hasta 10'"
    sql = sql & "   WHEN nNroEmp > 10 THEN '02. Más de 10' END"
    sql = sql & " From"
    sql = sql & " ("
    sql = sql & "   Select Convert(Varchar(7),dfecvig,102) Fecha, cCtaCod cuenta, nMontoDesemb Desembolso, nsaldocap saldo, ISNULL(F.nnroemp,0) nnroemp"
    sql = sql & "        FROM " & psServConsol & "CreditoConsolTotal X"
    sql = sql & "        inner join   " & psServConsol & "FuenteingresoConsol F on X.cNumFuente = F.cNumFuente"
    sql = sql & "     WHERE nPrdEstado in ('2020','2021','2022','2050','2030','2031','2032') and substring(cLineaCred,1,2) = '02'"
    sql = sql & "       and DateDiff(mm,dfecvig,'" & Format(DateAdd("m", I, sFechaIni), "yyyy/mm/dd") & "')=0 ) C "  '= 0 And DateDiff(dd,dfecvig,'" & Format(pdFechaCierre, "mm/dd/yyyy") & "') >= 0 ) C"
    sql = sql & " ) C1 Group by Fecha, NroTRaba Order by Fecha, NroTRaba"
    Set rs = GetQuery(sql)
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
Next I


' COLOCACIONES POR MONTO

' ++++++++++++++++++++++
' ++++++++++++++++++++++

' Falta por Query Sector

' ++++++++++++++++++++++
' ++++++++++++++++++++++

For I = 0 To 5
    sql = " Select  Convert(Varchar(7),dfecvig,102) Fecha,"
    sql = sql & " cDestCre = CASE WHEN nDestCre = '2' THEN 'A' WHEN nDestCre = '1' THEN 'B'"
    sql = sql & " WHEN nDestCre = '3' THEN 'C' END,"
    sql = sql & " COUNT(C.cCtaCod) numero, sum(nMontoDesemb) Desembolso"
    sql = sql & "       FROM " & psServConsol & "CreditoConsolTotal C"
    sql = sql & " WHERE nPrdEstado in ('2020','2021','2022','2050','2030','2031','2032') and substring(cLineaCred,1,2) = '02'"
    sql = sql & "     And DateDiff(dd,dFecVig,'" & Format(DateAdd("m", I, sFechaIni), "yyyy/mm/dd") & "') = 0 " 'And DateDiff(dd,dfecvig,'" & Format(pdFechaCierre, "mm/dd/yyyy") & "') >= 0"
    sql = sql & " GROUP BY Convert(Varchar(7),dfecvig,102), nDestCre ORDER BY Fecha, nDestCre"
    Set rs = GetQuery(sql)
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
Next I



'  COLOCACIONES POR MONTO
For I = 0 To 5
    sql = "Select Fecha, Rango, Count(Cuenta) Num, isnull(Sum(Desembolso),0) Desembolso, Sum(Saldo) Saldo FROM"
    sql = sql & " (  Select Fecha, C.cuenta, C.Desembolso, C.Saldo, Rango = CASE"
    sql = sql & "   WHEN C.Desembolso <= 2000 THEN '01. Hasta 2000 NS'"
    sql = sql & "   WHEN C.Desembolso > 2000 AND C.Desembolso <= 5000 THEN '02. De 2001 A 5000'"
    sql = sql & "   WHEN C.Desembolso > 5000 AND C.Desembolso <= 10000 THEN '03. De 5001 a 10000'"
    sql = sql & "   WHEN C.Desembolso > 10000  THEN '04. Mas de 10000' END"
    sql = sql & " From"
    sql = sql & " ("
    sql = sql & "   Select Convert(Varchar(7),dfecvig,102) Fecha, cCtaCod cuenta, nMontoDesemb Desembolso, nsaldocap saldo"
    sql = sql & "        FROM " & psServConsol & "CreditoConsolTotal WHERE nPrdEstado in ('2020','2021','2022','2050','2030','2031','2032') and substring(cLineaCred,1,2) = '02'"
    sql = sql & "       And DateDiff(dd,dFecVig,'" & Format(DateAdd("m", I, sFechaIni), "yyyy/mm/dd") & "') = 0) C " ' And DateDiff(dd,dfecvig,'2003/03/30') >= 0 ) C
    sql = sql & " ) C1 Group by Fecha, Rango Order by Fecha, Rango"
    Set rs = GetQuery(sql)
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
Next I

' Suma de las operaciones
For I = Asc("C") To Asc("H")
    xlHoja1.Range(Chr(I) & "19").Formula = "=" & Chr(I) & "20 +" & Chr(I) & "21 +" & Chr(I) & "22"
    xlHoja1.Range(Chr(I) & "24").Formula = "=" & Chr(I) & "25 +" & Chr(I) & "26 +" & Chr(I) & "27 +" & Chr(I) & "28"
    xlHoja1.Range(Chr(I) & "30").Formula = "=" & Chr(I) & "31 +" & Chr(I) & "32"
    xlHoja1.Range(Chr(I) & "34").Formula = "=" & Chr(I) & "35 +" & Chr(I) & "37"
    xlHoja1.Range(Chr(I) & "39").Formula = "=" & Chr(I) & "40 +" & Chr(I) & "41"
    xlHoja1.Range(Chr(I) & "42").Formula = "=" & Chr(I) & "43 +" & Chr(I) & "44"
    xlHoja1.Range(Chr(I) & "46").Formula = "=" & Chr(I) & "47 +" & Chr(I) & "48 +" & Chr(I) & "49 +" & Chr(I) & "50 +" & Chr(I) & "51"
    xlHoja1.Range(Chr(I) & "52").Formula = "=" & Chr(I) & "53 +" & Chr(I) & "54 +" & Chr(I) & "55 +" & Chr(I) & "56 +" & Chr(I) & "57"
    xlHoja1.Range(Chr(I) & "59").Formula = "=" & Chr(I) & "60 +" & Chr(I) & "61"
    xlHoja1.Range(Chr(I) & "62").Formula = "=" & Chr(I) & "64 +" & Chr(I) & "65"
    xlHoja1.Range(Chr(I) & "67").Formula = "=" & Chr(I) & "68 +" & Chr(I) & "69 +" & Chr(I) & "70 +" & Chr(I) & "71"
    xlHoja1.Range(Chr(I) & "72").Formula = "=" & Chr(I) & "73 +" & Chr(I) & "74 +" & Chr(I) & "75 +" & Chr(I) & "76"
Next I

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

Public Function nRepo179201_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
 sql = "   Select cPersCod,cPersNom,cCalifica as Total " _
            & " from " & psServerConsol & "RCDvc" & psFecha & "01 where (cCalifica ='' or cCalifica is null or cCalifica not in('0','1','2','3','4'))"

Set rs = GetQuery(sql)
Set nRepo179201_ = rs
Set rs = Nothing
End Function
Public Function nRepo179202_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "SELECT cPersCod,cPersNom, cTipPers, cCodSBS from " & psServerConsol & "RCDvc" & psFecha & "01  " _
        & " WHERE (PATINDEX('%Y/O%', cPersNom )<>0  OR  PATINDEX('% O %', cPersNom )<>0 OR PATINDEX('% Y %', cPersNom)<>0 )  and cTipPers<>'3' " _
        & " ORDER BY cTipPers"
Set rs = GetQuery(sql)
Set nRepo179202_ = rs
Set rs = Nothing
End Function
Public Function nRepo179203_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
 sql = "SELECT cPersCod,cCodSBS, cPersNom from " & psServerConsol & "RCDvc" & psFecha & "01  " _
        & " WHERE (cCodSbs IS NULL OR cCodSbs ='' or LEN(cCodSbs)<>10) "
Set rs = GetQuery(sql)
Set nRepo179203_ = rs
Set rs = Nothing
End Function
Public Function nRepo179204_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
 sql = "SELECT cPersCod,cPersNom,cTipPers,cMagEmp,cNudoTr FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE CTIPPERS ='2' AND CMAGEMP  NOT IN ('1','2','3','4') "
Set rs = GetQuery(sql)
Set nRepo179204_ = rs
Set rs = Nothing
End Function
Public Function nRepo179205_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
 sql = "SELECT cPersCod,cPersNom,cTipPers,cMagEmp,cNudoci FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE CTIPPERS NOT IN ('2') AND CMAGEMP IN ('1','2','3','4')"
Set rs = GetQuery(sql)
Set nRepo179205_ = rs
Set rs = Nothing
End Function
Public Function nRepo179206_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
 sql = "SELECT cPersCod,cPersNom,cTipPers,cMagEmp , cNudoci, cNudotr  FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (CMAGEMP IS NULL  OR CMAGEMP ='' OR CMAGEMP NOT IN ('0','1','2','3','4'))"
Set rs = GetQuery(sql)
Set nRepo179206_ = rs
Set rs = Nothing
End Function

Public Function nRepo179207_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
 sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cTidoci, cNudoci, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01  WHERE cTidoci not in ('1','2','3','4','5') and cTipPers in ('1','3')"
Set rs = GetQuery(sql)
Set nRepo179207_ = rs
Set rs = Nothing
End Function
Public Function nRepo179208_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
 sql = "SELECT cPersCod,cPersNom,cTipPers, cNudoci, cCodSbS FROM " & psServerConsol & "RCDvc" & psFecha & "01  WHERE (cNudoCi like '%000000%' or cNudoci = '' or cNudoci is null)  and cTipPers in ('1','3')"
Set rs = GetQuery(sql)
Set nRepo179208_ = rs
Set rs = Nothing
End Function

Public Function nRepo179209_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
 sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cNudoTr,cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (cNudoTr like '%000000%' or cNudoTr = '' or cNudoTr is null)  and cTipPers in ('2')"
Set rs = GetQuery(sql)
Set nRepo179209_ = rs
Set rs = Nothing
End Function
Public Function nRepo179210_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cTidotr, cNudotr, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE cTidotr not in ('2','3') and cTipPers in ('2')"
Set rs = GetQuery(sql)
Set nRepo179210_ = rs
Set rs = Nothing
End Function
Public Function nRepo179211_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "SELECT cPersCod,cPersNom,cTipPers, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (cTipPers not in ('1','2','3','4') or cTipPers is null or cTipPers='')"
Set rs = GetQuery(sql)
Set nRepo179211_ = rs
Set rs = Nothing
End Function
Public Function nRepo179212_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "SELECT cPersCod,cPersNom,cTipPers, cCodSbs,cActEcon  FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE len(cActeCon)<>4 or cActecon like '%000%' or cActecon ='' or cActecon is Null"
Set rs = GetQuery(sql)
Set nRepo179212_ = rs
Set rs = Nothing
End Function
Public Function nRepo179213_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cNudoTr, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (cSiglas is Not Null or cSiglas<>'') "
Set rs = GetQuery(sql)
Set nRepo179213_ = rs
Set rs = Nothing
End Function
Public Function nRepo179214_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "SELECT cPersCod,cPersNom,cTipPers, cCodRegPub, cNudoTr, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (cCodRegPub like '%00000000000%' or cCodRegPub is null or cCodRegPub ='' ) and cTipPers ='2'"
Set rs = GetQuery(sql)
Set nRepo179214_ = rs
Set rs = Nothing
End Function

Public Function nRepo179215_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "SELECT cPersCod, cPersNom   FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE  CTIPPERS IN ('1') ORDER BY cPersCod "
Set rs = GetQuery(sql)
Set nRepo179215_ = rs
Set rs = Nothing
End Function
Public Function nRepo179216_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cNudoTr,cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (len(cNudoTr)<>11 and cTidoTr='2' )  and cTipPers in ('2')  "
Set rs = GetQuery(sql)
Set nRepo179216_ = rs
Set rs = Nothing
End Function
Public Function nRepo179217_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql As String
sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cNudoTr,cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (len(cNudoCi)<>8 and cTidoci='1' )  and cTipPers in ('1','3')  "
Set rs = GetQuery(sql)
Set nRepo179217_ = rs
Set rs = Nothing
End Function
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
Dim I As Integer
Dim J As Integer
'RaiseEvent ShowProgress
'On Error Resume Next
sql = "select * from ConstSistema where nConsSisCod=80"
Set rs = GetQuery(sql)
lsCodigo = IIf(IsNull(rs!nConsSisValor), "", rs!nConsSisValor)
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
For I = 3 To 38
    xlHoja1.Cells(J, 1) = I
    J = J + 1
Next I


J = 15
For I = 1 To 6
    sql = "Select count(*) as Total from " & psServerConsol & " IBM1 where cTipDoc='" & I & "'"
    Set rs = GetQuery(sql)
    xlHoja1.Cells(J, 3) = rs!Total
    J = J + 1
Next I
xlHoja1.Cells(15, 2) = "NUM DEU CON DOCUM T/1"
xlHoja1.Cells(16, 2) = "NUM DEU CON DOCUM T/2"
xlHoja1.Cells(17, 2) = "NUM DEU CON DOCUM T/3"
xlHoja1.Cells(18, 2) = "NUM DEU CON DOCUM T/4"
xlHoja1.Cells(19, 2) = "NUM DEU CON DOCUM T/5"
xlHoja1.Cells(20, 2) = "NUM DEU CON DOCUM T/6"
'RaiseEvent Progress(2, 20)
J = 21
For I = 1 To 2
    sql = "Select count(*) as Total from " & psServerConsol & " IBM1 where cTipPers='" & I & "'"
    Set rs = GetQuery(sql)
    xlHoja1.Cells(J, 3) = rs!Total
    J = J + 1
Next I
xlHoja1.Cells(21, 2) = "NUM DEU PERSONA NATURAL"
xlHoja1.Cells(22, 2) = "NUM DEU PERSONA JURIDICA"

'**************
sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen30MN>0"
Set rs = GetQuery(sql)
xlHoja1.Cells(23, 3) = rs!Total
xlHoja1.Cells(23, 2) = "NUM DEU CUO VENC < 30 MN"

sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen31MN>0"
Set rs = GetQuery(sql)
xlHoja1.Cells(24, 3) = rs!Total
xlHoja1.Cells(24, 2) = "NUM DEU CUO VENC > 30 MN"

sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen30ME>0"
Set rs = GetQuery(sql)
xlHoja1.Cells(25, 3) = rs!Total
xlHoja1.Cells(25, 2) = "NUM DEU CUO VENC < 30 ME"

sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen31Me>0"
Set rs = GetQuery(sql)
xlHoja1.Cells(26, 3) = rs!Total
xlHoja1.Cells(26, 2) = "NUM DEU CUO VENC > 30 ME"
'****************

J = 27
For I = 0 To 4
    sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='" & I & "'"
    Set rs = GetQuery(sql)
    xlHoja1.Cells(J, 3) = rs!Total
    J = J + 1
Next I
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

sql = "SELECT replace(convert(varchar(20),SUM(nNormMN)),'.','')  AS TOTAL From " & psServerConsol & "IBM1"
Set rs = GetQuery(sql)
xlHoja1.Cells(34, 3) = "=" & rs!Total
xlHoja1.Cells(34, 2) = "MN DEUDA DIRECTA VIGENTE"

xlHoja1.Cells(35, 3) = "0"
xlHoja1.Cells(35, 2) = "MN DD ARREND FINACIERO"

sql = "SELECT replace(convert(varchar(20),SUM(nRefMN)),'.','') AS TOTAL From " & psServerConsol & "IBM1"
Set rs = GetQuery(sql)
xlHoja1.Cells(36, 3) = "=" & rs!Total
xlHoja1.Cells(36, 2) = "MN DEUDA DIREC REFINANCIADA"

sql = "SELECT replace(convert(varchar(20),SUM(nVen30MN)),'.','') AS TOTAL From " & psServerConsol & "IBM1"
Set rs = GetQuery(sql)
xlHoja1.Cells(37, 3) = "=" & rs!Total
xlHoja1.Cells(37, 2) = "MN DEUDA DIREC VENC < 30"

sql = "SELECT replace(convert(varchar(20),SUM(nVen31MN)),'.','') AS TOTAL From " & psServerConsol & "IBM1"
Set rs = GetQuery(sql)
xlHoja1.Cells(38, 3) = "=" & rs!Total
xlHoja1.Cells(38, 2) = "MN DEUDA DIREC VENC > 30"
'RaiseEvent Progress(7, 20)
sql = "SELECT replace(convert(varchar(20),SUM(nCobJudMN)),'.','') AS TOTAL From " & psServerConsol & "IBM1"
Set rs = GetQuery(sql)
xlHoja1.Cells(39, 3) = "=" & rs!Total
xlHoja1.Cells(39, 2) = "MN DEUDA DIREC COB JUDICIAL"

xlHoja1.Cells(40, 2) = "MN DEUDA INDIRECTA"
xlHoja1.Cells(40, 3) = "0"
xlHoja1.Cells(41, 2) = "MN DEUDA AVALADA"
xlHoja1.Cells(41, 3) = "0"

sql = "SELECT replace(convert(varchar(20),SUM(nLinCredMN)),'.','') AS TOTAL From " & psServerConsol & "IBM1"
Set rs = GetQuery(sql)
xlHoja1.Cells(42, 3) = "=" & rs!Total
xlHoja1.Cells(42, 2) = "MN LINEA DE CREDITO"

sql = "SELECT replace(convert(varchar(20),SUM(nCastMN)),'.','') AS TOTAL From " & psServerConsol & "IBM1"
Set rs = GetQuery(sql)
xlHoja1.Cells(43, 3) = "=" & rs!Total
xlHoja1.Cells(43, 2) = "MN CREDITOS CASTIGADOS"

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
For I = 34 To 50
    xlHoja1.Cells(I, 5) = J
    J = J + 1
Next I

'*************************
'RaiseEvent Progress(9, 20)

sql = "SELECT replace(convert(varchar(20),SUM(nNormME)),'.','') AS TOTAL From " & psServerConsol & "IBM1"
Set rs = GetQuery(sql)
xlHoja1.Cells(34, 7) = rs!Total
xlHoja1.Cells(34, 6) = "ME DEUDA DIRECTA VIGENTE"

xlHoja1.Cells(35, 7) = "0"
xlHoja1.Cells(35, 6) = "ME DD ARREND FINANCIERO"

sql = "SELECT replace(convert(varchar(20),SUM(nRefME)),'.','') AS TOTAL From " & psServerConsol & "IBM1"
Set rs = GetQuery(sql)
xlHoja1.Cells(36, 7) = rs!Total
xlHoja1.Cells(36, 6) = "ME DEUDA DIREC REFINANCIADA"

sql = "SELECT replace(convert(varchar(20),SUM(nVen30ME)),'.','') AS TOTAL From " & psServerConsol & "IBM1"
Set rs = GetQuery(sql)
xlHoja1.Cells(37, 7) = rs!Total
xlHoja1.Cells(37, 6) = "ME DEUDA DIREC VENC < 30"

sql = "SELECT replace(convert(varchar(20),SUM(nVen31ME)),'.','') AS TOTAL From " & psServerConsol & "IBM1"
Set rs = GetQuery(sql)
xlHoja1.Cells(38, 7) = rs!Total
xlHoja1.Cells(38, 6) = "ME DEUDA DIREC VENC > 30"

sql = "SELECT replace(convert(varchar(20),SUM(nCobJudME)),'.','') AS TOTAL From " & psServerConsol & "IBM1"
Set rs = GetQuery(sql)
xlHoja1.Cells(39, 7) = rs!Total
xlHoja1.Cells(39, 6) = "ME DEUDA DIREC COB JUDICIAL"

xlHoja1.Cells(40, 7) = "0"
xlHoja1.Cells(40, 6) = "ME DEUDA INDIRECTA"
'RaiseEvent Progress(11, 20)
xlHoja1.Cells(41, 7) = "0"
xlHoja1.Cells(41, 6) = "ME DEUDA LAVADA"

sql = "SELECT replace(convert(varchar(20),SUM(nLinCredME)),'.','') AS TOTAL From " & psServerConsol & "IBM1"
Set rs = GetQuery(sql)
xlHoja1.Cells(42, 7) = rs!Total
xlHoja1.Cells(42, 6) = "ME DEUDA LINEA DE CREDITO"

sql = "SELECT replace(convert(varchar(20),SUM(nCastME)),'.','') AS TOTAL From " & psServerConsol & "IBM1"
Set rs = GetQuery(sql)
xlHoja1.Cells(43, 7) = rs!Total
xlHoja1.Cells(43, 6) = "ME DEUDA CREDITOS CASTIGADOS"

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

Public Sub nRepo179101_ArchivoRCD(ByVal psServerConsol As String, ByVal pdfecSis As Date)
Dim SQL1 As String
Dim Sql2 As String
Dim CadPrint() As String
Dim lsCadenaAux As String
Dim lsCadenaPrint As String
Dim TotalCadenas As Long
Dim contTotal As Long
Dim Cont As Long

Dim rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset

Dim NúmeroArchivo As Integer
Dim lsNombreArchivo  As String

Dim lsApellPat As String
            Dim lsApellMat As String
Dim lsPrimerNombre As String
Dim lsSegundoNombre As String
Dim lsApellCony As String
Dim lsNombrePers As String
Dim Pos1 As Long
Dim Pos2 As Long
Dim Pos3 As Long
Dim lsCadAux As String
Dim lsCadena As String
Dim lsTipPer As String
Dim lsGenero As String
Dim lsEstado As String
Dim lsCondCta As String
Dim lsCondDisp As String

'''
Dim J As Long
Dim lscTidoci As String
Dim lscNudoCi As String
Dim lscTidoTr As String
Dim lscNudotr As String
Dim lsActividad As String
Dim lsCodSbs As String
Dim lsCalificacion As String
Dim lsCodPers As String
NúmeroArchivo = FreeFile

Dim lsPersonaAct As String
Dim lnSaldo As Currency

Dim lsNumSecICC As String
Dim Conecta As DConecta
Set Conecta = New DConecta
Dim lsIndRCC As String
' Crea el nombre de archivo.

Dim lsCodOfiInfICC As String
Dim lsUbicaGeoICC As String
Dim lsCodDptoICC As String

Dim gsCodEmpInfICC As String
Dim gnMontoMinimoICC As Integer

SQL1 = "select * from constsistema where nConsSisCod=80"
Set rs1 = GetQuery(SQL1)
gsCodEmpInfICC = rs1!nConsSisValor

SQL1 = "select top 1 * from " & psServerConsol & "rcdparametro  order by  cMEs desc"
Set rs1 = GetQuery(SQL1)
gnMontoMinimoICC = rs1!nMontoMin

lsNombreArchivo = App.path & "\Spooler\RCD" & Format(pdfecSis, "yyyymm") & ".108"
Open lsNombreArchivo For Output As #NúmeroArchivo

'**** Registro CABECERA
Print #NúmeroArchivo, "0106" & "01" & FillNum(Trim(gsCodEmpInfICC), 5, "0") & Format(pdfecSis, "yyyymmdd") & "012" _
        & Space(15) & ImpreFormat(EliminaPunto(IIf(IsNull(gnMontoMinimoICC), 0, gnMontoMinimoICC)), 15, 0, False) & Chr(13) & Chr(10);

SQL1 = "SELECT * FROM " & psServerConsol & "sysobjects WHERE name = '" & "RCDvc" & Format(pdfecSis, "yyyymm") & "01'"
Conecta.AbreConexion
Set rs1 = Conecta.CargaRecordSet(SQL1)

If rs1.EOF And rs1.BOF Then
    Conecta.CierraConexion
    Close #NúmeroArchivo
    MsgBox "Tabla [RCDvc" & Format(pdfecSis, "yyyymm") & "01]  No existe, Por favor comuniquese con el Dpto de Sistemas", vbInformation, "Aviso"
    Exit Sub
End If
      
SQL1 = " Select COUNT(*) AS TOTAL From " & psServerConsol & "RCDvc" & Format(pdfecSis, "yyyymm") & "01 R1 " _
     & " Inner Join " & psServerConsol & "RCDvc" & Format(pdfecSis, "yyyymm") & "02 R2 ON R2.cPersCod=R1.cPersCod "

Set rs1 = Conecta.CargaRecordSet(SQL1)
Conecta.CierraConexion

If Not RSVacio(rs1) Then
    contTotal = IIf(IsNull(rs1!Total), 0, rs1!Total)
End If
rs1.Close
Set rs1 = Nothing

SQL1 = "Select R1.cTipoFor, R1.cTipoInf, R1.cNumSec, R1.cCodSBS, R1.cPersCod, R1.cActEcon, " _
    & "   R1.cCodRegPub, R1.cTidoTr, R1.cNudoTr, R1.cTiDoci, R1.cNuDoci, " _
    & "   R1.cTipPers, R1.cResid, R1.cCalifica, R1.cMagEmp, R1.cAccionista, R1.cRelInst,  " _
    & "   R1.cPaisNac, R1.cSiglas, R1.cPersNom,R1.cPersNomCom,R1.cPersGenero,R1.cPersEstado, R2.cTipoFor cTipoFor2, R2.cTipoInf cTipoInf2 , " _
    & "   R2.cCodAge, R2.cUbicGeo, R2.cCtaCnt, R2.cTipoCred, R2.nSaldo, R2.nCondDias,   " _
    & "   R2.cPersCod cCodPers2, CAPEPAT, CAPEMAT,CAPECAS,CNOMBRE1,CNOMBRE2, R1.cIndRCC " _
    & " From " & psServerConsol & "RCDvc" & Format(pdfecSis, "yyyymm") & "01 R1 " _
    & "      Inner Join " & psServerConsol & "RCDvc" & Format(pdfecSis, "yyyymm") & "02 R2 ON R2.cPersCod=R1.cPersCod " _
    & " ORDER BY CAPEPAT, CAPEMAT,CAPECAS,CNOMBRE1,CNOMBRE2 "

Conecta.AbreConexion
Set rs1 = Conecta.CargaRecordSet(SQL1)
Conecta.CierraConexion
Cont = 0: J = 0
'Screen.MousePointer = 11

If Not RSVacio(rs1) Then
    RaiseEvent ShowProgress
    Do While Not rs1.EOF
        J = J + 1
            
        If lsPersonaAct <> rs1!cPersCod Then

            Cont = Cont + 1
            lsNumSecICC = FillNum(Trim(Str(Cont)), 6, "0")
            
            'lsNombrePers = PstaNombre(CadDerecha(Trim(Replace(rs1!cPersNom, "-", "", , , vbTextCompare)), 80))
            'lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "/", "", , , vbTextCompare)), 80)
            'lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ",", "", , , vbTextCompare)), 80)
            'lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ".", " ", , , vbTextCompare)), 80)
            'lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "Ñ", "#", , , vbTextCompare)), 80)
            'lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "ñ", "#", , , vbTextCompare)), 80)
            '''''''''''''''
            ''''''INSERCION DEL 11/06/2004 PARA SEPARAR NOMBRES  - ENME

            lsTipPer = rs1!cTipPers
            lsApellPat = "": lsApellMat = "": lsPrimerNombre = "": lsApellCony = "": lsSegundoNombre = ""
            lsNombrePers = rs1!cPersNomCom
            If Len(Trim(lsNombrePers)) = 0 Then
                    lsNombrePers = Trim(lsNombrePers)
            Else
                If InStr(1, lsNombrePers, "Y/O", vbTextCompare) = 0 And InStr(1, lsNombrePers, " Y ", vbTextCompare) = 0 And InStr(1, lsNombrePers, " O ", vbTextCompare) = 0 Then
                   lsNombrePers = Trim(lsNombrePers)
                End If
            End If
            lsNombrePers = Trim(Replace(lsNombrePers, "-/", "", , , vbTextCompare))
            lsNombrePers = Trim(Replace(lsNombrePers, "-", "", , , vbTextCompare))
            lsNombrePers = Trim(Replace(lsNombrePers, ".", " ", , , vbTextCompare))
            lsNombrePers = Trim(Replace(lsNombrePers, "//", "/", , , vbTextCompare))
            
            '-****Separa Nombre de Clientes
'                If InStr(1, lsNombrePers, "Y/O", vbTextCompare) = 0 And InStr(1, lsNombrePers, " Y ", vbTextCompare) = 0 And InStr(1, lsNombrePers, " O ", vbTextCompare) = 0 Then
'                    If Trim(lsTipPer) = "1" Then
'                        Pos1 = InStr(1, lsNombrePers, "/", vbTextCompare)
'                        Pos2 = InStr(1, lsNombrePers, "\", vbTextCompare)
'                        Pos3 = InStr(1, lsNombrePers, ",", vbTextCompare)
'                        If Pos1 > 0 Then
'                            lsApellPat = Mid(lsNombrePers, 1, Pos1 - 1)
'                            lsCadAux = Mid(lsNombrePers, Pos1 + 1, Len(lsNombrePers))
'                            Pos1 = InStr(1, lsCadAux, "\", vbTextCompare)
'                            If Pos1 > 0 Then
'                                lsApellMat = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                lsCadAux = Mid(lsCadAux, Pos1 + 1, Len(lsCadAux))
'                                If Len(Trim(lsApellMat)) = 0 Then
'                                    Pos1 = InStr(1, lsCadAux, ",", vbTextCompare)
'                                    If Pos1 > 0 Then
'                                        lsApellCony = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                        lsCadAux = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                        If Len(lsCadAux) > 0 Then
'                                            Pos1 = InStr(1, lsCadAux, " ", vbTextCompare)
'                                            If Pos1 > 0 Then
'                                                lsPrimerNombre = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                                lsCadAux = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                                If Len(lsCadAux) > 0 Then
'                                                    lsSegundoNombre = lsCadAux
'                                                End If
'                                            Else
'                                                lsPrimerNombre = lsCadAux
'                                            End If
'
'                                        End If
'
'                                    Else
'                                        lsApellCony = lsCadAux
'                                    End If
'                                Else
'                                    Pos1 = InStr(1, lsCadAux, ",", vbTextCompare)
'                                    If Pos1 > 0 Then
'                                        lsApellCony = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                        lsCadAux = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                        If Len(lsCadAux) > 0 Then
'                                            Pos1 = InStr(1, lsCadAux, " ", vbTextCompare)
'                                            If Pos1 > 0 Then
'                                                lsPrimerNombre = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                                lsCadAux = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                                If Len(lsCadAux) > 0 Then
'                                                    lsSegundoNombre = lsCadAux
'                                                End If
'                                            Else
'                                                lsPrimerNombre = lsCadAux
'                                            End If
'
'                                        End If
'
'                                    Else
'                                        lsApellCony = lsCadAux
'                                    End If
'
'                                End If
'                            Else
'                                Pos1 = InStr(1, lsCadAux, ",", vbTextCompare)
'                                If Pos1 > 0 Then
'                                    lsApellMat = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                    lsCadAux = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                        If Len(lsCadAux) > 0 Then
'                                            Pos1 = InStr(1, lsCadAux, " ", vbTextCompare)
'                                            If Pos1 > 0 Then
'                                                lsPrimerNombre = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                                lsCadAux = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                                If Len(lsCadAux) > 0 Then
'                                                    lsSegundoNombre = lsCadAux
'                                                End If
'                                            Else
'                                                lsPrimerNombre = lsCadAux
'                                            End If
'                                        End If
'                                End If
'                            End If
'                       ElseIf Pos2 > 0 Then
'                            lsApellPat = Mid(lsNombrePers, 1, Pos2 - 1)
'                            lsCadAux = Mid(lsNombrePers, Pos2 + 1, Len(lsNombrePers))
'                            lsApellMat = ""
'                                If Len(Trim(lsApellMat)) = 0 Then
'                                    Pos2 = InStr(1, lsCadAux, ",", vbTextCompare)
'                                    If Pos2 > 0 Then
'                                        lsApellCony = Trim(Mid(lsCadAux, 1, Pos2 - 1))
'                                        lsCadAux = Trim(Mid(lsCadAux, Pos2 + 1, Len(lsCadAux)))
'                                        If Len(lsCadAux) > 0 Then
'                                            Pos2 = InStr(1, lsCadAux, " ", vbTextCompare)
'                                            If Pos2 > 0 Then
'                                                lsPrimerNombre = Trim(Mid(lsCadAux, 1, Pos2 - 1))
'                                                lsCadAux = Trim(Mid(lsCadAux, Pos2 + 1, Len(lsCadAux)))
'                                                If Len(lsCadAux) > 0 Then
'                                                    lsSegundoNombre = lsCadAux
'                                                End If
'                                            Else
'                                                lsPrimerNombre = lsCadAux
'                                            End If
'
'                                        End If
'
'                                    Else
'                                        lsApellCony = lsCadAux
'                                    End If
'                                End If
'                        ElseIf Pos3 > 0 Then
'                            lsApellPat = Mid(lsNombrePers, 1, Pos3 - 1)
'                            lsCadAux = Mid(lsNombrePers, Pos3 + 1, Len(lsNombrePers))
'                            If Len(lsCadAux) > 0 Then
'                                 Pos3 = InStr(1, lsCadAux, " ", vbTextCompare)
'                                 If Pos3 > 0 Then
'                                     lsPrimerNombre = Trim(Mid(lsCadAux, 1, Pos3 - 1))
'                                     lsCadAux = Trim(Mid(lsCadAux, Pos3 + 1, Len(lsCadAux)))
'                                     If Len(lsCadAux) > 0 Then
'                                         lsSegundoNombre = lsCadAux
'                                     End If
'                                 Else
'                                     lsPrimerNombre = lsCadAux
'                                 End If
'
'                            End If
'                        End If
'                        ''''''''''''''''
'                        If Pos1 = 0 And Pos2 = 0 And Pos3 = 0 Then
'                            lsApellPat = lsNombrePers
'                            lsApellMat = ""
'                            lsPrimerNombre = ""
'                        End If
'                    Else
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "-", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "$", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "%", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "/", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ")", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "(", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "@", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "*", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "+", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "  ", " ", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "   ", " ", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "\", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ",", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ".", " ", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "Ñ", "#", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "ñ", "#", , , vbTextCompare)), 80)
'                        lsApellPat = lsNombrePers
'                        lsApellMat = ""
'                        lsPrimerNombre = ""
'                    End If
'                Else
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "-", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "$", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "%", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "/", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ")", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "(", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "@", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "*", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "+", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "  ", " ", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "   ", " ", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "\", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ",", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ".", " ", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "Ñ", "#", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "ñ", "#", , , vbTextCompare)), 80)
'                    lsApellPat = lsNombrePers
'                    lsApellMat = ""
'                    lsPrimerNombre = ""
'                    lsSegundoNombre = ""
'                End If
                
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Not IsNull(rs1!cPersGenero) Then
                lsGenero = Trim(rs1!cPersGenero)
            Else
                lsGenero = ""
            End If
            If Not IsNull(rs1!cPersEstado) Then
                Select Case Trim(rs1!cPersEstado)
                    Case "2", "5", "6"
                        lsEstado = "C"
                    Case "3"
                        lsEstado = "V"
                    Case "4"
                         lsEstado = "D"
                    Case "1"
                         lsEstado = "S"
                    Case Else
                         lsEstado = Trim(rs1!cPersEstado)
                End Select
            Else
                lsEstado = ""
            End If
                    
            ''''''''''''FIN DE INSERCION DEL 11/06/2004
                        
            Select Case Trim(rs1!cTipPers)
                Case "1", "3"
                    lscTidoci = IIf(IsNull(rs1!ctidoci), "", IIf(Len(Trim(rs1!ctidoci)) = 0, "", Trim(rs1!ctidoci)))
                    lscNudoCi = IIf(IsNull(rs1!cNudoci), "", IIf(Len(Trim(rs1!cNudoci)) = 0, "", Trim(rs1!cNudoci)))
                    lscTidoTr = ""
                    lscNudotr = ""
                Case Else
                    lscTidoci = ""
                    lscNudoCi = ""
                    lscTidoTr = IIf(IsNull(rs1!cTidoTr), "", IIf(Len(Trim(rs1!cTidoTr)) = 0, "", Trim(rs1!cTidoTr)))
                    lscNudotr = IIf(IsNull(rs1!cNudoTr), "", IIf(Len(Trim(rs1!cNudoTr)) = 0, "", Trim(rs1!cNudoTr)))
                    ' RUC DE 11 DIGITOS
                    lscTidoTr = IIf(lscTidoTr = "2", "3", lscTidoTr)
            End Select
            If Not IsNull(rs1!cCodSBS) Then
                lsCodSbs = Trim(rs1!cCodSBS)
            Else
                lsCodSbs = ""
            End If
            If Not IsNull(rs1!cActEcon) Then
                If Len(Trim(rs1!cActEcon)) > 0 Then
                    lsActividad = FillNum(rs1!cActEcon, 4, "0")
                Else
                    lsActividad = ""
                End If
            Else
                lsActividad = ""
            End If
            
            'lsCodPers = EmiteCodigoPersona(rs1!cPersCod, psServerConsol)
            'If Len(Trim(lsCodPers)) = 0 Then
                lsCodPers = Trim(rs1!cPersCod)
            'End If
            
            If lsCodPers = "1080800002739" Or lsCodPers = "1080100791482" Or lsCodPers = "1080100791482" Then
            '    Stop
            End If
            
            lsApellPat = Trim(rs1!CAPEPAT)
            lsApellMat = Trim(rs1!CAPEMAT)
            lsApellCony = Trim(rs1!CAPECAS)
            lsPrimerNombre = Trim(rs1!cNombre1)
            lsSegundoNombre = Trim(rs1!cNombre2)
            lsIndRCC = Trim(rs1!cIndRCC)
            
             Print #NúmeroArchivo, rs1!cTipoFor & rs1!cTipoInf & FillNum(Trim(lsNumSecICC), 8, " ") & _
                    ImpreFormat(IIf(Len(Trim(lsCodSbs)) = 0, "", lsCodSbs), 10, 0) & ImpreFormat(Trim(rs1!cPersCod), 20, 0) & FillNum(lsActividad, 4, " ") & _
                    FillNum(Trim(IIf(IsNull(rs1!ccodregpub), "", rs1!ccodregpub)), 15, " ") & _
                    FillNum(IIf(Len(Trim(lscTidoTr)) = 0, "", lscTidoTr), 1, " ") & _
                    FillNum(IIf(Len(Trim(lscNudotr)) = 0, "", lscNudotr), 11, " ") & FillNum(IIf(Len(Trim(lscTidoci)) = 0, "", lscTidoci), 1, " ") & _
                    ImpreFormat(IIf(Len(Trim(lscNudoCi)) = 0, "", lscNudoCi), 12, 0) & _
                    FillNum(Trim(IIf(IsNull(rs1!cTipPers), "", rs1!cTipPers)), 1, " ") & FillNum(Trim(IIf(IsNull(rs1!cResid), "", rs1!cResid)), 1, " ") & _
                    FillNum(Trim(rs1!cCalifica), 1, " ") & _
                    FillNum(Trim(IIf(IsNull(rs1!cMagEmp), "", rs1!cMagEmp)), 1, " ") & _
                    FillNum(Trim(IIf(IsNull(rs1!cAccionista), "", rs1!cAccionista)), 1, " ") & FillNum(Trim(IIf(IsNull(rs1!cRelInst), "", rs1!cRelInst)), 1, " ") & _
                    FillNum(Trim(rs1!cPaisNac), 4, " ") & FillNum(Trim(lsGenero), 1, " ") & FillNum(Trim(lsEstado), 1, " ") & _
                    ImpreFormat(Trim(IIf(IsNull(rs1!cSiglas), "", rs1!cSiglas)), 20, 0) & _
                    ImpreFormat(Trim(lsApellPat), 120, 0) & _
                    ImpreFormat(Trim(lsApellMat), 40, 0) & _
                    ImpreFormat(Trim(lsApellCony), 40, 0) & _
                    ImpreFormat(Trim(lsPrimerNombre), 40, 0) & _
                    ImpreFormat(Trim(lsSegundoNombre), 40, 0) & ImpreFormat(Trim(lsIndRCC), 1, 0) & Chr(13) & Chr(10);
         
        End If
        
        Print #NúmeroArchivo, rs1!cTipoFor2 & rs1!cTipoInf2 & FillNum(Trim(lsNumSecICC), 8, " ") & _
                FillNum(Trim(rs1!CCodage), 4, "0") & FillNum(Trim(rs1!cUbicGeo), 6, " ") & _
                FormatoCtaContable(rs1!cctacnt) & FillNum(Trim(rs1!cTipoCred), 1, " ") & _
                ImpreFormat(EliminaPunto(IIf(IsNull(rs1!nSaldo), 0, rs1!nSaldo)), 15, 0, False) & _
                FillNum(Trim(rs1!nCondDias), 4, " ") & "06" & "01" & Chr(13) & Chr(10);
        
        '********************* ACTUALIZANDO NUMERO DE SECUENCIA PARA REPORTE ICC *******************
        SQL1 = "Update " & psServerConsol & "RCDvc" & Format(pdfecSis, "yyyymm") & "01 SET CNUMSEC='" & Trim(lsNumSecICC) & "' WHERE cPersCod='" & Trim(rs1!cPersCod) & "'"
        Sql2 = "Update " & psServerConsol & "RCDvc" & Format(pdfecSis, "yyyymm") & "02 SET CNUMSEC='" & Trim(lsNumSecICC) & "' WHERE cPersCod='" & Trim(rs1!cPersCod) & "'"
        'Conecta.AbreConexion
        'Conecta.Ejecutar (SQL1)
        'Conecta.Ejecutar (Sql2)
        'Conecta.CierraConexion
        '*******************************************************************************************
        lsPersonaAct = rs1!cPersCod
        RaiseEvent Progress(rs1.Bookmark, rs1.RecordCount)
        rs1.MoveNext
    Loop
    RaiseEvent CloseProgress
End If
rs1.Close
Set rs1 = Nothing
Conecta.CierraConexion
'***************************************
'** Totales de la Empresa
'***************************************

Cont = Cont + 1
lsNumSecICC = FillNum(Trim(Str(Cont)), 6, "0")

Print #NúmeroArchivo, "2" & "1" & FillNum(Trim(lsNumSecICC), 6, " ") & _
        Space(10) & Chr(13) & Chr(10);

SQL1 = "Select * from " & psServerConsol & "RCDvc" & Format(pdfecSis, "yyyymm") & "03 Order by cCtaCnt,cTipoCred,nCondDias "

Conecta.AbreConexion
Set rs1 = Conecta.CargaRecordSet(SQL1)
Conecta.CierraConexion
If Not RSVacio(rs1) Then
    Do While Not rs1.EOF
        
        Print #NúmeroArchivo, "2" & "2" & FillNum(Trim(lsNumSecICC), 6, " ") & _
                FillNum(Trim(lsCodOfiInfICC), 4, " ") & FillNum(Trim(lsUbicaGeoICC), 6, " ") & FormatoCtaContable(rs1!cctacnt) & FillNum(Trim(rs1!cTipoCred), 1, " ") & _
                ImpreFormat(EliminaPunto(IIf(IsNull(rs1!nSaldo), 0, rs1!nSaldo)), 15, 0, False) & _
                FillNum(Trim(rs1!nCondDias), 4, " ") & Chr(13) & Chr(10);
        
        rs1.MoveNext
    Loop
End If
rs1.Close
Set rs1 = Nothing

Close #NúmeroArchivo   ' Cierra el archivo.
MsgBox "Se ha generado el Archivo RCDvc00.108 Satisfactoriamente, Termino : " & Time()
End Sub

Public Function FormatoCtaContable(ByVal pCuentaCnt As String) As String
  FormatoCtaContable = Trim(pCuentaCnt) & String(14 - Len(Trim(pCuentaCnt)), "0")
End Function

Public Function EmiteCodigoPersona(ByVal lsCodAux As String, ByVal psServerConsol As String) As String
Dim Co As DConecta
Dim sql As String
Dim rs As New ADODB.Recordset
Set Co = New DConecta
sql = "Select * from " & psServerConsol & "RCDCodigoAux where cCodAux='" & Trim(lsCodAux) & "'"
Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
Co.CierraConexion

If Not RSVacio(rs) Then
    EmiteCodigoPersona = Trim(rs!cPersCod)
Else
    EmiteCodigoPersona = ""
End If
rs.Close
Set rs = Nothing
Set Co = Nothing
End Function

Private Function EliminaPunto(lnNumero As Currency) As Currency
Dim Pos As Long
Dim CadAux As String
Dim CadAux1 As String
Dim lsNumero As String
lsNumero = Trim(Str(lnNumero))
If Val(lsNumero) > 0 Then
    Pos = InStr(1, lsNumero, ".", vbTextCompare)
    If Pos > 0 Then
        CadAux = Mid(lsNumero, 1, Pos - 1)
        CadAux1 = Mid(lsNumero, Pos + 1, Len(Trim(lsNumero)))
        If Len(Trim(CadAux1)) = 1 Then
            CadAux1 = CadAux1 & "0"
        End If
        EliminaPunto = CCur(CadAux & CadAux1)
    Else
        EliminaPunto = lnNumero & "00"
    End If
Else
    EliminaPunto = lnNumero
End If
End Function

Private Function EmiteCalificacion(lsCodPers As String) As String
Dim Co As DConecta
Dim sql As String
Dim rs As New ADODB.Recordset
Set Co = New DConecta
sql = "Select * from ColocevalCalif where cperscod='" & lsCodPers & "'"
Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
Co.CierraConexion
EmiteCalificacion = ""
If Not RSVacio(rs) Then
    EmiteCalificacion = Trim(rs!cEvalCalif)
End If
rs.Close
Set rs = Nothing
End Function

Public Sub nRepo179103_IBM(ByVal psServerConsol As String, ByVal dFecSis As Date, ByVal psNomCmact As String, ByVal pscodcmac As String)
Dim NúmeroArchivo As Integer
Dim lsNombreArchivo  As String
Dim lsNombrePers As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim contTotal  As Long
Dim sql As String
Dim J As Long
Dim lsApellPat As String
Dim lsApellMat As String
Dim lsApellCony As String
Dim lsNombre As String
Dim Pos As Long
Dim Pos1 As Long
Dim Pos2 As Long
Dim Cont As Long
Dim lsTipoDoc As String
Dim lsNumDoc As String
Dim lsPrefijos(6) As String
Dim I As Long
Dim m As Long
Dim lsCodPers As String
Dim lsCadAux As String

Dim TOTALTIPODOC1 As Long: Dim TOTALTIPODOC2 As Long: Dim TOTALTIPODOC3 As Long
Dim TOTALTIPODOC4 As Long: Dim TOTALTIPODOC5 As Long: Dim TOTALTIPODOC6 As Long

Dim TOTALPERSNAT As Long: Dim TOTALPERSJUR As Long

Dim TOTALVENCMEN30MN As Long: Dim TOTALVENCMAY30MN As Long
Dim TOTALVENCMEN30ME As Long: Dim TOTALVENCMAY30ME As Long

Dim TOTALCAL0 As Long: Dim TOTALCAL1 As Long: Dim TOTALCAL2 As Long:
Dim TOTALCAL3 As Long: Dim TOTALCAL4 As Long

Dim TOTALNORMMN As Currency: Dim TOTALREFMN As Currency: Dim TOTALMEN30MN As Currency
Dim TOTALMAY30MN  As Currency: Dim TOTALCOBJUDMN  As Currency: Dim TOTALLINCREDMN As Currency:
Dim TOTALCASTMN As Currency
Dim lsCalificacion As String
Dim TOTALNORMME As Currency: Dim TOTALREFME As Currency: Dim TOTALMEN30ME As Currency
Dim TOTALMAY30ME  As Currency: Dim TOTALCOBJUDME  As Currency:  Dim TOTALLINCREDME As Currency
Dim TOTALCASTME As Currency
'------->
Dim gsNombreEntInfIBM As String
Dim gsNomCortoEntInfIBM As String
Dim gsCodEmpInfICC As String
Dim lsNumSecICC As String
sql = "select nConsSisValor from ConstSistema where nConsSisCod=80"
Set rs = GetQuery(sql)
gsCodEmpInfICC = IIf(IsNull(rs!nConsSisValor), "", rs!nConsSisValor)
gsNombreEntInfIBM = psNomCmact
gsNomCortoEntInfIBM = pscodcmac
lsPrefijos(1) = "DE LA"
lsPrefijos(2) = "DE LOS"
lsPrefijos(3) = "DE "
lsPrefijos(4) = "LA "
lsPrefijos(5) = "DEL "
lsPrefijos(6) = "DI "

NúmeroArchivo = FreeFile
' Crea el nombre de archivo.

lsNombreArchivo = App.path & "\Spooler\ICS01.112"
Open lsNombreArchivo For Output As #NúmeroArchivo
Print #NúmeroArchivo, FillNum(Trim(gsCodEmpInfICC), 3, " ") & FillNum("1", 1, " ") & FillNum(Trim(Format(dFecSis, "YYYYMMDD")), 8, " ") & _
        Space(11) & ImpreFormat(gsNombreEntInfIBM, 80, 0) & ImpreFormat(Trim(gsNomCortoEntInfIBM), 10, 0) & _
        Space(546) & "X"
'------>

sql = "Select count(cCodPers) as Total From " & psServerConsol & "IBM1 where SUBSTRING(cApellPat,1,1) NOT IN('Ñ') "
Set rs = GetQuery(sql)


contTotal = rs!Total
rs.Close
Set rs = Nothing

sql = "Select cCodPers,cApellPat,cApellMat,cNombre,cTipPers,cTipDoc,cNumDoc,cCalifica," _
& " nNormMN,nRefMN,nVen30MN,nVen31MN,nCobJudMN,nLinCredMN,nCastMN,nNormME,nRefME, " _
& " nVen30ME , nVen31ME, nCobJudME, nLinCredME, nCastME " _
& " From " & psServerConsol & "IBM1 where SUBSTRING(cApellPat,1,1) <> 'Ñ' order by cApellPat,cApellmat "

'Select * From " & psServerConsol & "IBM1 where SUBSTRING(cApellPat,1,1) not in('Ñ') order by cApellPat,cApellmat"

Set rs = GetQuery(sql)
Cont = 0
If Not RSVacio(rs) Then
    Do While Not rs.EOF
        Cont = Cont + 1
        lsNumSecICC = FillNum(Trim(Str(Cont)), 6, "0")
        
        lsCalificacion = Trim(rs!cCalifica)
        'lsCalificacion = EmiteCalificacion(rs!cCodPers)
        'If Len(Trim(lsCalificacion)) = 0 Then
        '    lsCalificacion = Trim(rs!cCalifica)
        'End If
        
        Print #NúmeroArchivo, FillNum(Trim(gsCodEmpInfICC), 3, " ") & FillNum(Trim("2"), 1, " ") & FillNum(Trim(lsNumSecICC), 6, " ") & _
                    FillNum(Trim(rs!CTIPDOC), 1, " ") & ImpreFormat(Trim(rs!cNumDoc), 12, 0) & _
                    ImpreFormat(Trim(Replace(rs!cApellPat, "Ñ", "#")), 60, 0) & ImpreFormat(Trim(Replace(rs!cApellMat, "Ñ", "#")), 20, 0) & _
                    ImpreFormat(Trim(Replace(rs!cNombre, "Ñ", "#")), 30, 0) & FillNum(Trim(rs!cTipPers), 1, " ") & _
                    ImpreFormat(lsCalificacion, 1, 0) & Space(5) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nNormMN), 0, rs!nNormMN)), 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nRefMN), 0, rs!nRefMN)), 15, 0, False) & ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen30MN), 0, rs!nVen30MN)), 15, 0, False) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen31MN), 0, rs!nVen31MN)), 15, 0, False) & ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCobJudMN), 0, rs!nCobJudMN)), 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nLinCredMN), 0, rs!nLinCredMN)), 15, 0, False) & ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCastMN), 0, rs!nCastMN)), 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nNormME), 0, rs!nNormME)), 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nRefME), 0, rs!nRefME)), 15, 0, False) & ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen30ME), 0, rs!nVen30ME)), 15, 0, False) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen31ME), 0, rs!nVen31ME)), 15, 0, False) & ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCobJudME), 0, rs!nCobJudME)), 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nLinCredME), 0, rs!nLinCredME)), 15, 0, False) & ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCastME), 0, rs!nCastME)), 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & Space(9) & "X" & Chr(10);
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

sql = "Select count(cCodPers) as Total From " & psServerConsol & "IBM1 where SUBSTRING(cApellPat,1,1) IN('Ñ') "
Set rs = GetQuery(sql)

contTotal = rs!Total
rs.Close
Set rs = Nothing

sql = "Select * From " & psServerConsol & "IBM1 where SUBSTRING(CAPELLPAT,1,1) in('Ñ') order by cApellPat,cApellmat"
Set rs = GetQuery(sql)

m = 0
If Not RSVacio(rs) Then
    Do While Not rs.EOF
        m = m + 1
        Cont = Cont + 1
        lsNumSecICC = FillNum(Trim(Str(Cont)), 6, "0")
        lsCalificacion = Trim(rs!cCalifica)
        'lsCalificacion = EmiteCalificacion(rs!cCodPers)
        'If Len(Trim(lsCalificacion)) = 0 Then
        '    lsCalificacion = Trim(rs!cCalifica)
        'End If
        Print #NúmeroArchivo, FillNum(Trim(gsCodEmpInfICC), 3, " ") & FillNum(Trim("2"), 1, " ") & FillNum(Trim(lsNumSecICC), 6, " ") & _
                    FillNum(Trim(rs!CTIPDOC), 1, " ") & ImpreFormat(Trim(rs!cNumDoc), 12, 0) & _
                    ImpreFormat(Trim(Replace(rs!cApellPat, "Ñ", "#")), 60, 0) & ImpreFormat(Trim(Replace(rs!cApellMat, "Ñ", "#")), 20, 0) & _
                    ImpreFormat(Trim(Replace(rs!cNombre, "Ñ", "#")), 30, 0) & FillNum(Trim(rs!cTipPers), 1, " ") & _
                    ImpreFormat(lsCalificacion, 1, 0) & Space(5) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nNormMN), 0, rs!nNormMN)), 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nRefMN), 0, rs!nRefMN)), 15, 0, False) & ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen30MN), 0, rs!nVen30MN)), 15, 0, False) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen31MN), 0, rs!nVen31MN)), 15, 0, False) & ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCobJudMN), 0, rs!nCobJudMN)), 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nLinCredMN), 0, rs!nLinCredMN)), 15, 0, False) & ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCastMN), 0, rs!nCastMN)), 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nNormME), 0, rs!nNormME)), 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nRefME), 0, rs!nRefME)), 15, 0, False) & ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen30ME), 0, rs!nVen30ME)), 15, 0, False) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen31ME), 0, rs!nVen31ME)), 15, 0, False) & ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCobJudME), 0, rs!nCobJudME)), 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(EliminaPunto(IIf(IsNull(rs!nLinCredME), 0, rs!nLinCredME)), 15, 0, False) & ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCastME), 0, rs!nCastME)), 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
                    ImpreFormat(0, 15, 0, False) & Space(9) & "X" & Chr(10);
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

sql = " SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPPERS='1'"
Set rs1 = GetQuery(sql)
TOTALPERSNAT = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPPERS='2'"
Set rs1 = GetQuery(sql)
TOTALPERSJUR = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close


sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='1'"
Set rs1 = GetQuery(sql)
TOTALTIPODOC1 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='2'"
Set rs1 = GetQuery(sql)
TOTALTIPODOC2 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='3'"
Set rs1 = GetQuery(sql)
TOTALTIPODOC3 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='4'"
Set rs1 = GetQuery(sql)
TOTALTIPODOC4 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='5'"
Set rs1 = GetQuery(sql)
TOTALTIPODOC5 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='6'"
Set rs1 = GetQuery(sql)
TOTALTIPODOC6 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
TOTALVENCMEN30MN = 0
sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen30MN>0"
Set rs1 = GetQuery(sql)
TOTALVENCMEN30MN = TOTALVENCMEN30MN + IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

TOTALVENCMEN30ME = 0
sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen30ME>0"
Set rs1 = GetQuery(sql)
TOTALVENCMEN30ME = TOTALVENCMEN30ME + IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

TOTALVENCMAY30MN = 0
sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen31MN>0"
Set rs1 = GetQuery(sql)
TOTALVENCMAY30MN = TOTALVENCMAY30MN + IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

TOTALVENCMAY30ME = 0
sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen31Me>0"
Set rs1 = GetQuery(sql)
TOTALVENCMAY30ME = TOTALVENCMAY30ME + IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='0'"
Set rs1 = GetQuery(sql)
TOTALCAL0 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='1'"
Set rs1 = GetQuery(sql)
TOTALCAL1 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='2'"
Set rs1 = GetQuery(sql)
TOTALCAL2 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='3'"
Set rs1 = GetQuery(sql)
TOTALCAL3 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='4'"
Set rs1 = GetQuery(sql)
TOTALCAL4 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

sql = "SELECT SUM(nNormMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(sql)
TOTALNORMMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
sql = "SELECT SUM(nRefMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(sql)
TOTALREFMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
sql = "SELECT SUM(nVen30MN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(sql)
TOTALMEN30MN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
sql = "SELECT SUM(nVen31MN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(sql)
TOTALMAY30MN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
sql = "SELECT SUM(nCobJudMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(sql)
TOTALCOBJUDMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
sql = "SELECT SUM(nLinCredMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(sql)
TOTALLINCREDMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
sql = "SELECT SUM(nCastMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(sql)
TOTALCASTMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close

'monto soles
sql = "SELECT SUM(nNormME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(sql)
TOTALNORMME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
sql = "SELECT SUM(nRefME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(sql)
TOTALREFME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
sql = "SELECT SUM(nVen30ME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(sql)
TOTALMEN30ME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
sql = "SELECT SUM(nVen31ME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(sql)
TOTALMAY30ME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
sql = "SELECT SUM(nCobJudME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(sql)
TOTALCOBJUDME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
sql = "SELECT SUM(nLinCredME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(sql)
TOTALLINCREDME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
sql = "SELECT SUM(nCastME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(sql)
TOTALCASTME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close

Print #NúmeroArchivo, FillNum(Trim(gsCodEmpInfICC), 3, " ") & FillNum(Trim("3"), 1, " ") & _
ImpreFormat(TOTALTIPODOC1, 8, 0) & ImpreFormat(TOTALTIPODOC2, 8, 0) & ImpreFormat(TOTALTIPODOC3, 8, 0) & ImpreFormat(TOTALTIPODOC4, 8, 0) & ImpreFormat(TOTALTIPODOC5, 8, 0) & ImpreFormat(TOTALTIPODOC6, 8, 0) & _
ImpreFormat(TOTALPERSNAT, 8, 0) & ImpreFormat(TOTALPERSJUR, 8, 0) & _
ImpreFormat(TOTALVENCMEN30MN, 8, 0) & ImpreFormat(TOTALVENCMAY30MN, 8, 0) & ImpreFormat(TOTALVENCMEN30ME, 8, 0) & ImpreFormat(TOTALVENCMAY30ME, 8, 0) & _
ImpreFormat(TOTALCAL0, 8, 0) & ImpreFormat(TOTALCAL1, 8, 0) & ImpreFormat(TOTALCAL2, 8, 0) & ImpreFormat(TOTALCAL3, 8, 0) & ImpreFormat(TOTALCAL4, 8, 0) & _
ImpreFormat(EliminaPunto(TOTALNORMMN), 15, 0, False) & ImpreFormat(0, 15, 0, False) & ImpreFormat(EliminaPunto(TOTALREFMN), 15, 0, False) & _
ImpreFormat(EliminaPunto(TOTALMEN30MN), 15, 0, False) & ImpreFormat(EliminaPunto(TOTALMAY30MN), 15, 0, False) & ImpreFormat(EliminaPunto(TOTALCOBJUDMN), 15, 0, False) & _
ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & ImpreFormat(EliminaPunto(TOTALLINCREDMN), 15, 0, False) & ImpreFormat(EliminaPunto(TOTALCASTMN), 15, 0, False) & _
ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
ImpreFormat(0, 15, 0, False) & _
ImpreFormat(EliminaPunto(TOTALNORMME), 15, 0, False) & ImpreFormat(0, 15, 0, False) & ImpreFormat(EliminaPunto(TOTALREFME), 15, 0, False) & _
ImpreFormat(EliminaPunto(TOTALMEN30ME), 15, 0, False) & ImpreFormat(EliminaPunto(TOTALMAY30ME), 15, 0, False) & ImpreFormat(EliminaPunto(TOTALCOBJUDME), 15, 0, False) & _
ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & ImpreFormat(EliminaPunto(TOTALLINCREDME), 15, 0, False) & ImpreFormat(EliminaPunto(TOTALCASTME), 15, 0, False) & _
ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & ImpreFormat(0, 15, 0, False) & _
ImpreFormat(0, 15, 0, False) & Space(9) & "X" & Chr(10);


Close NúmeroArchivo
MsgBox "Reporte Generado Satisfactoriamente en " & lsNombreArchivo, vbInformation, "Aviso"
End Sub

Private Sub Class_Initialize()
Set Conecta = New DConecta
End Sub


Private Sub Class_Terminate()
Set Conecta = Nothing
End Sub

Private Function ValidaTabla(pdfecSis As Date) As Boolean
Dim sql As String
Dim rs As Recordset

sql = "SELECT * FROM sysobjects WHERE name = '" & "RCDvc" & Format(pdfecSis, "yyyymm") & "01'"
Conecta.AbreConexion
Set rs = Conecta.CargaRecordSet(sql)
ValidaTabla = True

If rs.EOF And rs.BOF Then
    Conecta.CierraConexion
    MsgBox "Tabla [RCDvc" & Format(pdfecSis, "yyyymm") & "01]  No existe, Por favor comuniquese con el Dpto de Sistemas", vbInformation, "Aviso"
    ValidaTabla = False
End If

End Function


