Attribute VB_Name = "gReportes"
Option Explicit

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim lbLibroOpen As Boolean
Dim lsArchivo As String
Dim xlHoja1 As Excel.Worksheet
Dim nLin As Long

Public Sub ReporteTarjetasPorAgencia(ByVal pnCodAge As Integer)
Dim P As Previo.clsPrevio
Dim R As ADODB.Recordset
Dim sSQL As String
Dim sCadRep As String
Dim Cont As Integer
Dim loConec As New DConecta

    
    sSQL = " REP_TarjetasPorAgencia " & pnCodAge
    
    Set R = New ADODB.Recordset
    sCadRep = "."
    
    'Cabecera
    sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
    sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Space(42) & "Usuario : " & gsCodUser & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(40) & "REPORTE DE TARJETAS POR AGENCIA" & Chr(10) & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
    sCadRep = sCadRep & Space(5) & "TARJETA" & Space(20) & "ESTADO" & Chr(10)
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
    Cont = 0
    
    'AbrirConexion
    loConec.AbreConexion
    R.Open sSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not R.EOF
        sCadRep = sCadRep & Space(5) & Right(Space(16) & R!cNumTarjeta, 16) & Space(5) & R!cDescrip & Chr(10)
        Cont = Cont + 1
        R.MoveNext
    Loop
    R.Close
    'CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    Set R = Nothing
    
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
    sCadRep = sCadRep & Space(5) & "CANTIDAD : " & Str(Cont) & Chr(10)
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
    
        Set P = New Previo.clsPrevio
        Call P.Show(sCadRep, "REPORTE")
        Set P = Nothing
    
End Sub

Public Sub ReporteTarjetasEmitidaPorAgencia(ByVal pnCodAge As Integer, ByVal psAgencia As String)
Dim P As Previo.clsPrevio
Dim R As ADODB.Recordset
Dim sSQL As String
Dim sCadRep As String
Dim Cont As Integer
Dim ContTotal As Integer
Dim nCodAge As Integer
Dim loConec As New DConecta

    sSQL = " REP_TarjetasEmitidasPorAgencia " & pnCodAge
    
    Set R = New ADODB.Recordset
    sCadRep = "."
    
    'Cabecera
    sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
    sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Space(42) & "Usuario : " & gsCodUser & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(40) & "REPORTE DE TARJETAS EMITIDAS POR AGENCIA" & Chr(10) & Chr(10)

    Cont = 0
    ContTotal = 0
    'AbrirConexion
    loConec.AbreConexion
    R.Open sSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    If Not R.EOF Then
        nCodAge = R!nCodAge
            sCadRep = sCadRep & Space(5) & "AGENCIA : " & R!cNomAgeArea & Chr(10) & Chr(10) & Chr(10)
            sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
            sCadRep = sCadRep & Space(5) & "TARJETA" & Space(20) & "ESTADO" & Chr(10)
            sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
    End If
    Do While Not R.EOF
        sCadRep = sCadRep & Space(5) & Right(Space(16) & R!cNumTarjeta, 16) & Space(5) & R!cDescrip & Chr(10)
        Cont = Cont + 1
        ContTotal = ContTotal + 1
        R.MoveNext
        If Not R.EOF Then
            
            If nCodAge <> R!nCodAge Then
                    nCodAge = R!nCodAge
                    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
                    sCadRep = sCadRep & Space(5) & "CANTIDAD : " & Str(Cont) & Chr(10)
                    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10) & Chr(10)
    
                    sCadRep = sCadRep & Space(5) & "AGENCIA : " & R!cNomAgeArea & Chr(10) & Chr(10) & Chr(10)
                    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
                    sCadRep = sCadRep & Space(5) & "TARJETA" & Space(20) & "ESTADO" & Chr(10)
                    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
                    
                    Cont = 0
            End If
        Else
                    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
                    sCadRep = sCadRep & Space(5) & "CANTIDAD : " & Str(Cont) & Chr(10)
                    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10) & Chr(10)
        
        End If
    Loop
    R.Close
    'CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    Set R = Nothing
    
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
    sCadRep = sCadRep & Space(5) & "CANTIDAD TOTAL : " & Str(ContTotal) & Chr(10)
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
    
        Set P = New Previo.clsPrevio
        Call P.Show(sCadRep, "REPORTE")
        Set P = Nothing
    
End Sub


Public Sub ListadoDERemesasENTransito()
Dim P As Previo.clsPrevio
Dim R As ADODB.Recordset
Dim sSQL As String
Dim sCadRep As String
Dim nTotal As Double
Dim loConec As New DConecta


sSQL = " ATM_RepRemesasENTransito "

Set R = New ADODB.Recordset
sCadRep = "."

'Cabecera
sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Space(42) & "Usuario : " & gsCodUser & Chr(10) & Chr(10)
sCadRep = sCadRep & Space(40) & "LISTADO DE REMESAS EN TRANSITO" & Chr(10) & Chr(10) & Chr(10)
sCadRep = sCadRep & Space(5) & String(120, "-") & Chr(10)
sCadRep = sCadRep & Space(5) & "FECHA" & Space(10) & "DESCRIPCION" & Space(10) & "ORIGEN" & Space(15) & "DESTINO" & Space(8) & "NUM. TARJ. INICIAL" & Space(2) & "NUM. TARJ. FINAL" & Space(2) & "CANTIDAD" & Chr(10)
sCadRep = sCadRep & Space(5) & String(120, "-") & Chr(10)

nTotal = 0

'AbrirConexion
loConec.AbreConexion
R.Open sSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Do While Not R.EOF
    sCadRep = sCadRep & Space(5) & Format(R!dFecha, "dd/mm/yyyy") & Space(2) & Left(R!cDesc & Space(20), 20) & Left(R!cOrigen & Space(20), 20) & Left(R!cDestino & Space(20), 20) & Left(R!cNumInicial & Space(20), 20) & Left(R!cNumFinal & Space(20), 20) & Right(Space(5) & Format(R!nCantidad, "#0"), 5) & Chr(10)
        nTotal = nTotal + 1
    R.MoveNext
Loop
R.Close
'CerrarConexion
loConec.CierraConexion
Set loConec = Nothing
Set R = Nothing

sCadRep = sCadRep & Space(5) & String(120, "-") & Chr(10)
sCadRep = sCadRep & Space(5) & "NUMERO DE REGISTROS : " & Right(Space(5) & Format(nTotal, "#0"), 5) & Chr(10)
sCadRep = sCadRep & Space(5) & String(120, "-") & Chr(10)


    Set P = New Previo.clsPrevio
    Call P.Show(sCadRep, "REPORTE")
    Set P = Nothing


End Sub


Public Sub ListadoDETarjetasRetiradas()
Dim P As Previo.clsPrevio
Dim R As ADODB.Recordset
Dim sSQL As String
Dim sCadRep As String
Dim Cont As Integer
Dim ContTotal As Integer
Dim nCodAge As Integer
Dim loConec As New DConecta

    sSQL = " ATM_RepTarjetasRetiradas "
    
    Set R = New ADODB.Recordset
    sCadRep = "."
    
    'Cabecera
    sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
    sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Space(42) & "Usuario : " & gsCodUser & Chr(10) & Chr(10)
    sCadRep = sCadRep & Space(40) & "REPORTE DE TARJETAS RETIRADAS POR AGENCIA" & Chr(10) & Chr(10)

    Cont = 0
    ContTotal = 0
    'AbrirConexion
    loConec.AbreConexion
    R.Open sSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    If Not R.EOF Then
        nCodAge = R!nCodAge
            sCadRep = sCadRep & Space(5) & "AGENCIA : " & R!cNomAgeArea & Chr(10) & Chr(10) & Chr(10)
            sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
            sCadRep = sCadRep & Space(5) & "TARJETA" & Chr(10)
            sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
    End If
    Do While Not R.EOF
        sCadRep = sCadRep & Space(5) & Right(Space(16) & R!cNumTarjeta, 16) & Chr(10)
        Cont = Cont + 1
        ContTotal = ContTotal + 1
        R.MoveNext
        If Not R.EOF Then
            
            If nCodAge <> R!nCodAge Then
                    nCodAge = R!nCodAge
                    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
                    sCadRep = sCadRep & Space(5) & "CANTIDAD : " & Str(Cont) & Chr(10)
                    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10) & Chr(10)
    
                    sCadRep = sCadRep & Space(5) & "AGENCIA : " & R!cNomAgeArea & Chr(10) & Chr(10) & Chr(10)
                    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)

                    
                    Cont = 0
            End If
        Else
                    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
                    sCadRep = sCadRep & Space(5) & "CANTIDAD : " & Str(Cont) & Chr(10)
                    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10) & Chr(10)
        
        End If
    Loop
    R.Close
    'CerrarConexion
    loConec.CierraConexion
    Set loConec = Nothing
    Set R = Nothing
    
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
    sCadRep = sCadRep & Space(5) & "CANTIDAD TOTAL : " & Str(ContTotal) & Chr(10)
    sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
    
        Set P = New Previo.clsPrevio
        Call P.Show(sCadRep, "REPORTE")
        Set P = Nothing
    
End Sub

Public Sub ListadoDERemesasENTransitoFueraDELimite()

Dim P As Previo.clsPrevio
Dim R As ADODB.Recordset
Dim sSQL As String
Dim sCadRep As String
Dim nTotal As Double
Dim loConec As New DConecta

sSQL = " ATM_RepRemesasENTransitoFueraLim "

Set R = New ADODB.Recordset
sCadRep = "."

'Cabecera
sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Space(42) & "Usuario : " & gsCodUser & Chr(10) & Chr(10)
sCadRep = sCadRep & Space(40) & "LISTADO DE REMESAS EN TRANSITO FUERA DE LIMITE" & Chr(10) & Chr(10) & Chr(10)
sCadRep = sCadRep & Space(5) & String(120, "-") & Chr(10)
sCadRep = sCadRep & Space(5) & "FECHA" & Space(10) & "DESCRIPCION" & Space(10) & "ORIGEN" & Space(15) & "DESTINO" & Space(8) & "NUM. TARJ. INICIAL" & Space(2) & "NUM. TARJ. FINAL" & Space(2) & "CANTIDAD" & Chr(10)
sCadRep = sCadRep & Space(5) & String(120, "-") & Chr(10)

nTotal = 0

'AbrirConexion
loConec.AbreConexion
R.Open sSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Do While Not R.EOF
    sCadRep = sCadRep & Space(5) & Format(R!dFecha, "dd/mm/yyyy") & Space(2) & Left(R!cDesc & Space(20), 20) & Left(R!cOrigen & Space(20), 20) & Left(R!cDestino & Space(20), 20) & Left(R!cNumInicial & Space(20), 20) & Left(R!cNumFinal & Space(20), 20) & Right(Space(5) & Format(R!nCantidad, "#0"), 5) & Chr(10)
        nTotal = nTotal + 1
    R.MoveNext
Loop
R.Close
'CerrarConexion
loConec.CierraConexion
Set loConec = Nothing
Set R = Nothing

sCadRep = sCadRep & Space(5) & String(120, "-") & Chr(10)
sCadRep = sCadRep & Space(5) & "NUMERO DE REGISTROS : " & Right(Space(5) & Format(nTotal, "#0"), 5) & Chr(10)
sCadRep = sCadRep & Space(5) & String(120, "-") & Chr(10)


    Set P = New Previo.clsPrevio
    Call P.Show(sCadRep, "REPORTE")
    Set P = Nothing


End Sub



Public Sub ListadoDEEstadisticasRegStockBGEN()

Dim P As Previo.clsPrevio
Dim R As ADODB.Recordset
Dim sSQL As String
Dim sCadRep As String
Dim nTotal As Double
Dim loConec As New DConecta

sSQL = " REP_EstaRegStockBGEN "

Set R = New ADODB.Recordset
sCadRep = "."

'Cabecera
sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Space(42) & "Usuario : " & gsCodUser & Chr(10) & Chr(10)
sCadRep = sCadRep & Space(40) & "ESTADISTICAS DE REGISTRO DE STOCKS DE BOVEDA GENEARAL" & Chr(10) & Chr(10) & Chr(10)
sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
sCadRep = sCadRep & Space(15) & "FECHA" & Space(10) & "CANTIDAD" & Chr(10)
sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)

nTotal = 0

'AbrirConexion
loConec.AbreConexion
R.Open sSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Do While Not R.EOF
    sCadRep = sCadRep & Space(5) & Format(R!dFecha, "dd/mm/yyyy") & Space(5) & Right(Space(30) & Format(R!nCantidad, "#0.00"), 16) & Chr(10)
        nTotal = nTotal + R!nCantidad
    R.MoveNext
Loop
R.Close
'CerrarConexion
loConec.CierraConexion
Set loConec = Nothing
Set R = Nothing

sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
sCadRep = sCadRep & Space(21) & Space(15) & Space(5) & Right(Space(30) & Format(nTotal, "#0.00"), 16) & Chr(10)
sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)


    Set P = New Previo.clsPrevio
    Call P.Show(sCadRep, "REPORTE")
    Set P = Nothing


End Sub

Public Sub ListadoStocksPorReporte()

Dim P As Previo.clsPrevio
Dim R As ADODB.Recordset
Dim sSQL As String
Dim sCadRep As String
Dim nTotal As Double
Dim loConec As New DConecta

sSQL = " REP_Stocks_PorCajero " & Trim(Str(CInt(gsCodAge)))

Set R = New ADODB.Recordset
sCadRep = "."

'Cabecera
sCadRep = sCadRep & Space(5) & "CMAC MAYNAS S.A." & Space(50) & "FECHA : " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & Chr(10)
sCadRep = sCadRep & Space(5) & "SIMACC-Tarjeta de Debito" & Space(42) & "Usuario : " & gsCodUser & Chr(10) & Chr(10)
sCadRep = sCadRep & Space(40) & "REPORTE DE STOCKS DE TARJETAS POR CAJERO" & Chr(10) & Chr(10) & Chr(10)
sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
sCadRep = sCadRep & Space(15) & "USUARIO" & Space(4) & "FECHA REGISTRO" & Space(9) & "CANTIDAD" & Chr(10)
sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)

nTotal = 0

'AbrirConexion
loConec.AbreConexion
R.Open sSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Do While Not R.EOF
    sCadRep = sCadRep & Space(5) & Right(Space(16) & R!cCodUsu, 16) & Space(5) & Format(R!dFecha, "dd/mm/yyyy") & Space(5) & Right(Space(30) & Format(R!nCantidad, "#0.00"), 16) & Chr(10)
        nTotal = nTotal + R!nCantidad
    R.MoveNext
Loop
R.Close
'CerrarConexion
loConec.CierraConexion
Set loConec = Nothing
Set R = Nothing

sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)
sCadRep = sCadRep & Space(21) & Space(15) & Space(5) & Right(Space(30) & Format(nTotal, "#0.00"), 16) & Chr(10)
sCadRep = sCadRep & Space(5) & String(100, "-") & Chr(10)


    Set P = New Previo.clsPrevio
    Call P.Show(sCadRep, "REPORTE")
    Set P = Nothing


End Sub

Public Sub ListadoReporteControlOpe(ByVal pdFecIni As Date, ByVal pdFecFin As Date)

Dim R As ADODB.Recordset
Dim sSQL As String
Dim sCadRep As String
Dim nTotal As Double
Dim loConec As New DConecta
Dim lsHoja As String
Dim lsFecIni As String
Dim lsFecFin As String

lsFecIni = Format(Trim(Str(pdFecIni)), "yyyy-mm-dd")
lsFecFin = Format(Trim(Str(pdFecFin)), "yyyy-mm-dd")

sSQL = " REP_CONTROL_OPERACIONES " & "'" & lsFecIni & "','" & lsFecFin & "'"

Set R = New ADODB.Recordset

loConec.AbreConexion
R.Open sSQL, loConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText

If R.RecordCount = 0 Then
    MsgBox "No existen datos para el reporte", vbExclamation, "Mensaje del Sistema"
    Exit Sub
End If

lsArchivo = App.Path & "\SPOOLER\RepControlOpeATM_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
If Not lbLibroOpen Then
    Exit Sub
End If
nLin = 1

lsHoja = "Control_Operaciones"
ExcelAddHoja lsHoja, xlLibro, xlHoja1

xlHoja1.Range("A1:S1").EntireColumn.Font.FontStyle = "Arial"
xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.CenterHorizontally = True
xlHoja1.PageSetup.Zoom = 75
xlHoja1.PageSetup.TopMargin = 2

xlHoja1.Range("A1:A1").RowHeight = 17
xlHoja1.Range("A1:A1").ColumnWidth = 14 'STATUS
xlHoja1.Range("B1:B1").ColumnWidth = 14 'Cod Rspta
xlHoja1.Range("C1:C1").ColumnWidth = 14 'Origen MSG
xlHoja1.Range("D1:D1").ColumnWidth = 14  'Origen Tran
xlHoja1.Range("E1:E1").ColumnWidth = 14 'Motivo Ext
xlHoja1.Range("F1:F1").ColumnWidth = 14 'Total

xlHoja1.Cells(nLin, 2) = "Reporte de control de Operaciones ATM"
xlHoja1.Range("A" & nLin & ":F" & nLin).Merge True
xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
xlHoja1.Range("A" & nLin & ":F" & nLin).HorizontalAlignment = xlHAlignCenter
nLin = nLin + 1
xlHoja1.Cells(nLin, 2) = "Desde " & pdFecIni & " Hasta " & pdFecFin
xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
xlHoja1.Range("A" & nLin & ":F" & nLin).Merge True
xlHoja1.Range("A" & nLin & ":F" & nLin).HorizontalAlignment = xlHAlignCenter
    
nLin = nLin + 2

xlHoja1.Cells(nLin, 1) = "STATUS"
xlHoja1.Cells(nLin, 2) = "COD Respuesta"
xlHoja1.Cells(nLin, 3) = "Origen MSG"
xlHoja1.Cells(nLin, 4) = "Origen Tran"
xlHoja1.Cells(nLin, 5) = "Motivo Extorno"
xlHoja1.Cells(nLin, 6) = "Total"

xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
xlHoja1.Range("A" & nLin & ":F" & nLin).HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("A" & nLin & ":F" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
xlHoja1.Range("A" & nLin & ":F" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A" & nLin & ":F" & nLin).Borders(xlInsideVertical).Color = vbBlack
xlHoja1.Range("A" & nLin & ":F" & nLin).Interior.Color = 13421619

With xlHoja1.PageSetup
    .LeftHeader = ""
    .CenterHeader = ""
    .RightHeader = ""
    .LeftFooter = ""
    .CenterFooter = ""
    .RightFooter = ""

    .PrintHeadings = False
    .PrintGridlines = False
    .PrintComments = xlPrintNoComments
    .CenterHorizontally = True
    .CenterVertically = False
    .Orientation = xlLandscape
    .Draft = False
    .FirstPageNumber = xlAutomatic
    .Order = xlDownThenOver
    .BlackAndWhite = False
    .Zoom = 55
End With

nLin = nLin + 1

Do While Not R.EOF
    xlHoja1.Range("A" & nLin & ":E" & nLin).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Cells(nLin, 1) = R!STATUS_MSG
    xlHoja1.Cells(nLin, 2) = R!COD_RESP
    xlHoja1.Cells(nLin, 3) = R!ORIGEN_MSG
    xlHoja1.Cells(nLin, 4) = R!ORIGEN_TRAN
    xlHoja1.Cells(nLin, 5) = R!IND_MOTIVO_EXT
    xlHoja1.Cells(nLin, 6) = R!Total
    
    nLin = nLin + 1
    R.MoveNext
    If R.EOF Then
        Exit Do
    End If
Loop
R.Close
'CerrarConexion
loConec.CierraConexion
Set loConec = Nothing
Set R = Nothing

ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
CargaArchivo lsArchivo, App.Path & "\SPOOLER\"

End Sub







