VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DReportesRiesgos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
Public Event ShowProgress()
Public Event CloseProgress()
Public Event Progress(pnValor As Long, pnTotal As Long)

Dim ServerCons As String
Dim RCD As nRcdReportes


Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet  'Micro
Dim fs As Scripting.FileSystemObject

Public Sub ReporteComercial()
    Dim rs As ADODB.Recordset
    Dim fs As New Scripting.FileSystemObject
    Dim nFila As Integer
    Dim nCol As Integer
    Dim nTotal As Integer
    Dim I As Integer
    Dim lbExisteHoja As Boolean
    Dim objConecta As DConecta
    Dim sArchivo As String
    Dim bExcel As Boolean
    Dim dFecha As Date
    Dim cUser As String
    Dim lsHoja As String
    Dim lsArchivo As String
    Dim sCodCta As String
    Dim sCodOficina As String
    Dim sCodProducto As String
    Dim sRelacion As String
    Dim sNombre As String
    Dim sCreditoRelacionado As String
    Dim sCreditoRefinanciado As String
    Dim nSaldoCapital As Double
    Dim nDiasAtraso As Integer
    Dim nCalifAnt As Integer
    Dim nCalifVig As Integer
    Dim sCodAnalista As String
    Dim nPrdRelacion As Integer
    Dim cFecVig As String
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    
    defecha = GetUltimaFechaCierre
    Set objConecta = New DConecta
    RaiseEvent ShowProgress
     cUser = gsCodUser
    RaiseEvent Progress(1, 10)
    lsHoja = CStr(dFecha) & " - " & "Comercial"
    RaiseEvent Progress(3, 10)
    lsArchivo = "COMERCIAL" & Format(dFecha, "ddmmyyyy") & ".XLS"
    Set fs = New Scripting.FileSystemObject
    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER" & lsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    lbExisteHoja = False
    
    For Each xlHoja1 In xlLibro.Worksheets
        If xlHoja1.Name = lsHoja Then
            xlHoja1.Activate
            xlHoja1.Range("A1", "AZ10000") = ""
            lbExisteHoja = True
            Exit For
        End If
    Next
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlLibro.Worksheets.Add
        xlHoja1.Name = lsNomHoja
    End If
    
    nFila = 2
    xlHoja1.PageSetup.Zoom = 80
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlAplicacion.Range("A1:R1000").Font.Size = 11
    
    
    ' configurando el tamaño de las columnas
    xlHoja1.Range("A1").ColumnWidth = 12 'codoficina
    xlHoja1.Range("B1").ColumnWidth = 12 'Producto
    xlHoja1.Range("C1").ColumnWidth = 12 'relacion con la cuenta
    xlHoja1.Range("D1").ColumnWidth = 25 ' descripcion de la relacion
    xlHoja1.Range("E1").ColumnWidth = 35 ' Descripcion del cliente
    xlHoja1.Range("F1").ColumnWidth = 30 'Credito relacionado
    xlHoja1.Range("G1").ColumnWidth = 30 ' Num Credito
    xlHoja1.Range("H1").ColumnWidth = 10 ' Normal/Refinanciado
    xlHoja1.Range("I1").ColumnWidth = 25 ' saldo capital
    xlHoja1.Range("J1").ColumnWidth = 15 ' Dias Atrasi CV
    xlHoja1.Range("K1").ColumnWidth = 15 ' Dias atrasado ca
    xlHoja1.Range("L1").ColumnWidth = 15 ' Calif Ant.
    xlHoja1.Range("M1").ColumnWidth = 15 ' Calif Act
    xlHoja1.Range("N1").ColumnWidth = 20 ' Fecha de vigencia
    xlHoja1.Range("O1").ColumnWidth = 10 ' estado
    xlHoja1.Range("Q1").ColumnWidth = 15 ' Codigo del analista
    
    
    'Colocando la cabecera
    xlHoja1.Cells(nFila, 1) = gsNomCmac
    xlHoja1.Cells(nFila, 7) = ""
    xlHoja1.Range(xlHoja1.Cells(nFila, 1), xlHoja1.Cells(nFila, 10)).Font.Bold = True
    
    nFila = nFila + 2
    xlHoja1.Cells(nfila1, 4) = "Reporte de Creditos ....al " & CStr(dFecha)
    xlHoja1.Range(xlHoja1.Cells(nFila, 1), xlHoja1.Cells(nFila, 10)).Font.Bold = True
    
    nFila = nFila + 2
    
    With xlHoja1
        .Cells(nFila, 1) = "Cod.Oficina"
        .Cells(nFila, 2) = "Producto"
        .Cells(nFila, 3) = "Relación"
        .Cells(nFila, 4) = "Descripcion Relacion"
        .Cells(nFila, 5) = "Nombre del Cliente"
        .Cells(nFila, 6) = "Crédito.Relac"
        .Cells(nFila, 7) = "Num.Crédito"
        .Cells(nFila, 8) = "Nor/Ref."
        .Cells(nFila, 9) = "Saldo Cap."
        .Cells(nFila, 10) = "Días Atra(CV)"
        .Cells(nFila, 11) = "Días Atra(CA)"
        .Cells(nFila, 12) = "Calif.Ant"
        .Cells(nFila, 13) = "Calif. Act."
        .Cells(nFila, 14) = "Fec. Vig."
        .Cells(nFila, 15) = "Estado"
        .Cells(nFila, 16) = "Cod.Ana"
        .Range(.Cells(nFila, 1), .Cells(nFila, 12)).Font.Bold = True
        .Range(.Cells(nFila, 1), .Cells(nFila, 16)).HorizontalAlignment = xlCenter
        .Range(.Cells(nFila, 1), .Cells(nFila, 16)).AutoFilter = True
    End With
    
    ' obteniendo el codigo de la oficina
        strSQL = " select  substring(a.cctacod,4,2) as CodOficina, "
        strSQL = strSQL & " Substring(a.cctacod,6,3) as Producto, "
        strSQL = strSQL & " p.nPrdPersRelac, "
        strSQL = strSQL & " Relacion=(Select cConsDescripcion from constante where p.nPrdPersRelac=nConsValor and nconscod='3002'), "
        strSQL = strSQL & " pp.cPersNombre as cPersNombre, "
        strSQL = strSQL & " p1.cctacod as CreditoRelac, "
        strSQL = strSQL & " a.cctacod  as NumCredito, "
        strSQL = strSQL & " a.cRefinan, "
        strSQL = strSQL & " a.nSaldoCap, "
        strSQL = strSQL & " a.nDiasAtraso, "
        strSQL = strSQL & " c.nDiasAtraso as nDiaAtraso1"
        strSQL = strSQL = " c.cCalGen as CalifAnt, "
        strSQL = strSQL & " a.cCalGen as CalifVig, "
        strSQL = strSQL & " a.dFecVig, "
        strSQL = strSQL & "a.nprdestado  as Estado"
        strSQL = strSQL & " a.cCodAnalista "
        strSQL = strSQL & " from dbcmacicaconsol..ColocCalifProv a "
        strSQL = strSQL & " inner join productopersona p on p.cctacod=a.cctacod "
        strSQL = strSQL & " inner join persona pp on pp.cperscod=p.cperscod "
        strSQL = strSQL & " inner join productopersona p1 on p1.cperscod=pp.cperscod "
        strSQL = strSQL & "inner join dbcmacicaconsol..ColocCalifProvTotal c on c.cperscod=a.cperscod and c.cctacod=a.cctacod and c.nprdestado=a.nprdestado "
        strSQL = strSQL & " where a.nprdestado in ('2020','2021','2022','2030','2031','2032','2201','2101','2104','2106','2107') and substring(a.cctacod,6,3) in ('101','102','103','121') "
        strSQL = strSQL & " and Datepart(mm,c.dfecvig)=Datepart(mm,getdate())-1"

    Dim nRelacion As Integer
    Dim cEstado As String
    Dim cRefina As String
    objConecta.AbreConexion
    Set rs = objConecta.CargaRecordSet(strSQL)
    RaiseEvent Progress(6, 10)
    nFila = nFila + 1
    Do Until rs.EOF
        sCodCta = IIf(IsNull(rs!cCtaCod), "", rs!cCtaCod)
        sCodOficina = IIf(IsNull(rs!CodOficina), "", rs!CodOficina)
        sCodProducto = IIf(IsNull(rs!Producto), "", rs!Producto)
        nRelacion = IIf(IsNull(rs!nPrdPersRelac), 0, rs!nPrdPersRelac)
        sRelacion = IIf(IsNull(rs!Relacion), "", rs!Relacion)
        sNombre = IIf(IsNull(rs!cPersNombre), "", rs!cPersNombre)
        sCreditoRelacionado = IIf(IsNull(rs!CreditoRelac), "", rs!CreditoRelac)
        sCreditoRefinanciado = IIf(IsNull(rs!NumCredito), "", rs!NumCredito)
        cRefina = IIf(IsNull(rs!cRefinan), "", rs!cRefinan)
        nSaldoCapital = IIf(IsNull(rs!nSaldoCap), 0, rs!nSaldoCap)
        nDiasAtraso1 = IIf(IsNull(rs!nDiasAtraso1), 0, rs!nDiasAtraso1)
        nDiasAtraso = IIf(IsNull(rs!nDiasAtraso), 0, rs!nDiasAtraso)
        nCalifAnt = IIf(IsNull(rs!CalifAnt), "", rs!CalifAnt)
        nCalifVig = IIf(IsNull(rs!CalifVig), "", rs!CalifVig)
        cFecVig = IIf(IsNull(rs!dFecVig), "", rs!dFecVig)
        cEstado = IIf(IsNull(rs!Estado), "", rs!Estado)
        sCodAnalista = IIf(IsNull(rs!cCodAnalista), "", rs!cCodAnalista)
        
        With xlHoja1
            .Cells(nFila, 1) = sCodOficina
            .Range(.Cells(nFila, 1), .Cells(nFila, 1)).HorizontalAlignment = xlCenter
            .Cells(nFila, 2) = sCodProducto
            .Range(.Cells(nFila, 2), .Cells(nFila, 2)).HorizontalAlignment = xlCenter
            .Cells(nFila, 3) = nRelacion
            .Range(.Cells(nFila, 3), .Cells(nFila, 3)).HorizontalAlignment = xlCenter
            .Cells(nFila, 4) = sRelacion
            .Range(.Cells(nFila, 4), .Cells(nFila, 4)).HorizontalAlignment = xlLeft
            .Cells(nFila, 5) = sNombre
            .Range(.Cells(nFila, 5), .Cells(nFila, 5)).HorizontalAlignment = xlLeft
            .Cells(nFila, 6) = sCreditoRelacionado
            .Range(.Cells(nFila, 6), .Cells(nFila, 6)).HorizontalAlignment = xlCenter
            .Cells(nFila, 7) = sCreditoRefinanciado
            .Range(.Cells(nFila, 7), .Cells(nFila, 7)).HorizontalAlignment = xlCenter
            .Cells(nFila, 8) = cRefina
            .Range(.Cells(nFila, 8), .Cells(nFila, 8)).HorizontalAlignment = xlCenter
            .Cells(nFila, 9) = Format(nSaldoCapital, "##0.00")
            .Range(.Cells(nFila, 9), .Cells(nFila, 7)).HorizontalAlignment = xlRight
            '.Cells(nFila, 10) = nDiasAtraso
            .Range(.Cells(nFila, 10), .Cells(nFila, 10)).HorizontalAlignment = xlRight
            SetearValoresNegativos nDiasAtraso, nFila, 10
            ' .Cells(nFila, 11) = nDiasAtraso1
            .Range(.Cells(nFila, 11), .Cells(nFila, 11)).HorizontalAlignment = xlRight
            SetearValoresNegativos ndiasaatraso1, nFila, 11
            .Cells(nFila, 12) = nCalifAnt
            .Range(.Cells(nFila, 12), .Cells(nFila, 12)).HorizontalAlignment = xlRight
            .Cells(nFila, 13) = nCalifVig
            .Range(.Cells(nFila, 13), .Cells(nFila, 13)).HorizontalAlignment = xlRight
            .Cells(nFila, 14) = cFecVig
            .Range(.Cells(nFila, 14), .Cells(nFila, 14)).HorizontalAlignment = xlRight
            .Cells(nFila, 15) = cEstado
            .Range(.Cells(nFila, 15), .Cells(nFila, 15)).HorizontalAlignment = xlCenter
            .Cells(nFila, 16) = sCodAnalista
            .Range(.Cells(nFila, 16), .Cells(nFila, 16)).HorizontalAlignment = xlCenter
            
        End With
        nFila = nFila + 1
        
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    objConecta.CierraConexion
    Set objConecta = Nothing
    xlHoja1.PageSetup.Zoom = 75
    RaiseEvent Progress(9, 10)
    xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
    xlLibro.Close
    xlAplicacion.Quit
    RaiseEvent Progress(10, 10)
    MsgBox "Se genero el archivo en" & App.path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xhoja1 = Nothing
    RaiseEvent CloseProgress
    Exit Sub
ErrHandler:
    If Not objConecta Is Nothing Then Set objConecta = Nothing
    If Not fs Is Nothing Then Set fs = Nothing
    MsgBox "Se ha producido un error en el sistema", vbInformation, "AVISO"
End Sub
Public Function GetUltimaFechaCierre() As Date
Dim loConecta As DConecta
Dim rs As ADODB.Recordset
Dim sql As String
sql = "select nConsSisValor from constsistema where nConsSisCod=14"
Set loConecta = New DConecta
Set rs = New ADODB.Recordset
loConecta.AbreConexion
Set rs = loConecta.CargaRecordSet(sql)
loConecta.CierraConexion
GetUltimaFechaCierre = rs!nConsSisValor
End Function


Public Sub Reporte_ClientesReclasificados()
 Dim rs As ADODB.Recordset
 Dim fs As New Scripting.FileSystemObject
 Dim lnFila As Integer
 Dim lnCol As Integer
 Dim lbExisteHoja As Boolean
 Dim lsNombreHoja As String
 Dim objConecta As DConecta
 Dim ldFechaCierre As Date
 Dim lbExcel As Boolean
 Dim lsCadena As String
 Dim cUser As String
 Dim lsArchivo As String
 Dim sSQL As String
 
 Dim cNombre As String
 Dim cCtaCod As String
 Dim cEstado As String
 Dim dFecVig As Date
 Dim cRefinac As String
 Dim nCuotasApr As Integer
 Dim nAnosVig As Integer
 Dim nMesVig As Integer
 Dim nMontoCol As Double
 Dim nSaldoCap As Double
 Dim nPorcentaje As Double
 Dim nDiasMaximo As Integer
 Dim nDiasAtraso As Integer
 Dim nDiasAtraso1 As Integer
 Dim ncCalifAnt As Integer
 Dim nCalifVig As Integer
 Dim nMoneda As Integer
 Dim nProducto As Integer
 Dim nAgencia As Integer
 Dim cPersCod As String
 'Dim nDiasMaximo As Integer
 Dim nCalif As Integer
    On Error GoTo ErrHandler
    ldFechaCierre = GetUltimaFechaCierre
    Set objConecta = New DConecta
    RaiseEvent ShowProgress
    cUser = gsCodUser
    lsNombreHoja = Format(ldFechaCierre, "ddmmyyyy") & "Clien.Refinan"
    RaiseEvent Progress(3, 10)
    lsArchivo = "ClienteRefina" & Format(ldFechaCierre, "ddmmyyyy") & ".XLS"
    Set fs = New Scripting.FileSystemObject
    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
        Set xlibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
    Else
        Set xlibro = xlAplicacion.Workbooks.Add
    End If
    lbExisteHoja = True
    For Each xlHoja1 In xlLibro.Worksheets
        If xlHoja1.Name = lsNombreHoja Then
            xlHoja1.Activate
            xlHoja1.Range("A1", "AZ0000") = ""
            lbExisteHoja = False
            Exit For
        End If
    Next
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlLibro.Worksheets.Add
        xlHoja1.Name = lsNombreHoja
    End If
    
    nFila = 2
    xlHoja1.PageSetup.Zoom = 80
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlAplicacion.Range("A1:R1000").Font.Size = 9
    
    ' configurando las columnas
    With xlHoja1
        .Range("A1").ColumnWidth = 50
        .Range("B1").ColumnWidth = 15
        .Range("C1").ColumnWidth = 15
        .Range("D1").ColumnWidth = 10
        .Range("E1").ColumnWidth = 15
        .Range("F1").ColumnWidth = 10
        .Range("G1").ColumnWidth = 15
        .Range("H1").ColumnWidth = 20
        .Range("I1").ColumnWidth = 20
        .Range("J1").ColumnWidth = 20
        .Range("K1").ColumnWidth = 25
        .Range("L1").ColumnWidth = 25
        .Range("M1").ColumnWidth = 25
        .Range("N1").ColumnWidth = 25
        .Range("O1").ColumnWidth = 25
        .Range("P1").ColumnWidth = 25
        .Range("Q1").ColumnWidth = 25
        .Range("R1").ColumnWidth = 25
        .Range("S1").ColumnWidth = 15
        .Range("T1").ColumnWidth = 15
        .Range("U1").ColumnWidth = 15
        
        .Cells(nFila, 1) = gsNomCmac
        .Cells(nFila, 7) = ""
        .Range(.Cells(nFila, 1), .Cells(nFila, 21)).Font.Bold = True
        
        nFila = nFila + 2
        .Cells(nFila, 4) = "Reporte de los Clientes Reclasificados al " & ldFechaCierre
        
        .Cells(nFila, 1) = "Nombre del Cliente"
        .Cells(nFila, 2) = "Cod.Cliente"
        .Cells(nFila, 3) = "Cod.Crédito"
        .Cells(nFila, 4) = "Estado"
        .Cells(nFila, 5) = "Fecv.Vigente"
        .Cells(nFila, 6) = "Nro/Ref."
        .Cells(nFila, 7) = "Cuotas Apr."
        .Cells(nFila, 8) = "Años Vig."
        .Cells(nFila, 9) = "Mes Vig."
        .Cells(nFila, 10) = "Cap.Desem"
        .Cells(nFila, 11) = "Saldo Cap."
        .Cells(nFila, 12) = "%Pagado"
        .Cells(nFila, 13) = "Dias Max"
        .Cells(nFila, 14) = "Dias Atr. (CV)"
        .Cells(nFila, 15) = "Días Atr. (CA)"
        .Cells(nFila, 16) = "Calif. Ant."
        .Cells(nFila, 17) = "Calif. Act."
        .Cells(nFila, 18) = "Calif."
        .Cells(nFila, 19) = "Moneda"
        .Cells(nFila, 20) = "Producto"
        .Cells(nFila, 21) = "Agencia"
        .Range(.Cells(nFila, 1), .Cells(nFila, 12)).Font.Bold = True
        .Range(.Cells(nFila, 1), .Cells(nFila, 16)).HorizontalAlignment = xlCenter
         .Range(.Cells(nFila, 1), .Cells(nFila, 16)).AutoFilter = True
         
        sSQL = "select c.cPersNombre as Nombre, "
        sSQL = ssq & "c.cPersCod as CodCliente, "
        sSQL = sSQL & " a.cCtacod as CodCredito, "
        sSQL = sSQL & "estado=(Select cConsDescripcion From Constante where nconscod='3001' and nConsValor=a.nPrdEstado), "
        sSQL = sSQL & "a.dFecVig, "
        sSQL = sslq & "a.cRefinan, "
        sSQL = sSQL & " CuotasApr=(Select a1.nCuotas "
        sSQL = sSQL & "from colocacestado a1 "
        sSQL = sSQL & " inner join producto b1 on a1.cCtaCod=b1.cCtaCod and a1.nPrdEstado=b.nPrdEstado "
        sSQL = sSQL & " where b1.nPrdEstado in ('2020','2021','2022','2030','2031','2032','2201','2101','2104','2106','2107') and "
        sSQL = sslq & "b1.cCtaCod=a.cCtaCod), "
        sSQL = sSQL & " nAnosVig=(Select Datediff(yyyy,nPrdEstado,getdate()) From colocacestado where cctacod=a.cctacod), "
        sSQL = sSQL & " nMesVig=(Select DateDiff(mm,nPrdEstado,getdate()) From colocacestado where cctacod=a.cctacod), "
        sSQL = sSQL & " d.nMontoCol,"
        sSQL = sSQL & " a.nSaldoCap,"
        sSQL = sSQL & " nDiasMaximo=(select Max(Datediff(dd,a1.dvenc,a1.dpago)) as DiasMaximo--,a.dvenc,a.dpago "
        sSQL = sSQL & "From ColocCalendario a1 "
        sSQL = sSQL & " inner join colocaccred b1 on a1.cctacod=b1.cctacod and a1.nnrocalen=b1.nnrocalen and a1.ncoloccalendapl=1 "
        sSQL = sSQL & " where a1.cctacod='108013041000000027' and a1.dpago is not null and a1.dvenc is not null), "
        sSQL = sSQL & " a.DiasAtraso, "
        sSQL = sSQL & " e.DiasAtraso as nDiasAtraso1, "
        sSQL = sSQL & " e.cCalGen as CalifAnt, "
        sSQL = sSQL & " a.cCalNor as CalifVig, "
        sSQL = sSQL & " a.cCalEval as Calif"
        sSQL = sSQL & " substring(a.cCtacod,9,1) as nMoneda, "
        sSQL = sSQL & " substring(a.cCtacod,6,3) as nProducto, "
        sSQL = sSQL & " substring(a.cCtaCod, 4, 2) As nAgencia "
        sSQL = sSQL & " from  dbcmacicaconsol..ColocCalifProv a "
        sSQL = sSQL & "inner join ProductoPersona b on a.cCtaCod=b.cCtaCod and nPrdPersRelac=20 "
        sSQL = sSQL & "inner join Persona c on c.cPersCod=b.cPersCod"
        sSQL = sSQL & "inner join Colocaciones d on d.cCtaCod=a.cCtaCod"
        sSQL = sSQL & " inner join dbcmacicaconsol..ColocCalifProvtotal e on e.cctacod=a.cctacod and Datepart(mm,e.dfecvig)=Datepart(mm,getdate())-1 "
        sSQL = sSQL & " where a.nprdestado in ('2020','2021','2022','2030','2031','2032','2201','2101','2104','2106','2107') and a.cref='R'"
        
        objConecta.AbreConexion
        Set rs = objConecta.CargaRecordSet(sSQL)
        objConecta.CierraConexion
        RaiseEvent ShowProgress
        nFila = nFila + 1
        Do While rs.EOF
           cNombre = IIf(IsNull(rs!Nombre), "", rs!Nombre)
           cPersCod = IIf(IsNull(rs!CodCliente), "", rs!cPersCod)
           cCtaCod = IIf(IsNull(rs!CodCredito), "", rs!CodCredito)
           cEstado = IIf(IsNull(rs!Estado), "", rs!Estado)
           dFecVig = IIf(IsNull(rs!dFecVig), Now, rs!dFecVig)
           cRefinac = IIf(IsNull(rs!cRefinan), "", rs1cRefinan)
           nCuotasApr = IIf(IsNull(rs!CuotasApr), 0, rs!CuotasApr)
           nAnosVig = IIf(IsNull(rs!nAnosVig), 0, rs!nAnosVig)
           nMesVig = IIf(IsNull(rs!nMesVig), 0, rs!nMesVig)
           nMontoCol = IIf(IsNull(rs!nMontoCol), 0, rs!nMontoCol)
           nSaldoCap = IIf(IsNull(rs!nSaldoCap), 0, rs!nSaldoCap)
           nPorcentaje = ((nMontoCol - nSaldoCap) / nMontoCol) * 100
           nDiasMaximo = IIf(IsNull(rs!nDiasMaximo), 0, rs!nDiasMaximo)
           nDiasAtraso = IIf(IsNull(rs!DiasAtraso), 0, rs!DiasAtraso)
           nDiasAtraso1 = IIf(IsNull(nDiasAtraso1), 0, rs!nDiasAtraso1)
           ncCalifAnt = IIf(IsNull(rs!CalifAnt), 0, rs!CalifAnt)
           nCalifVig = IIf(IsNull(rs!CalifVig), 0, rs!CalifVig)
           nCalif = IIf(IsNull(rs!Calif), 0, rs!Calif)
           nMoneda = IIf(IsNull(rs!nMoneda), 0, rs!nMoneda)
           nProducto = IIf(IsNull(rs!nProducto), 0, rs!nProducto)
           nAgencia = IIf(IsNull(rs!nAgencia), 0, rs!nAgencia)
           nDiasMaximo = IIf(IsNull(rs!nDiasMaximo), 0, rs!nDiasMaximo)
          .Cells(nFila, 1) = cNombre
          .Range(.Cells(nFila, 1), .Cells(nFila, 1)).HorizontalAlignment = xlLeft
          .Cells(nFila, 2) = cPersCod
          .Range(.Cells(nFila, 2), .Cells(nFila, 2)).HorizontalAlignment = xlCenter
          .Cells(nFila, 3) = cCtaCod
          .Range(.Cells(nFila, 3), .Cells(nFila, 3)).HorizontalAlignment = xlCenter
          .Cells(nFila, 4) = cEstado
          .Range(.Cells(nFila, 4), .Cells(nFila, 4)).HorizontalAlignment = xlCenter
          .Cells(nFila, 5) = dFecVig
          .Range(.Cells(nFila, 5), .Cells(nFila, 5)).HorizontalAlignment = xlCenter
          .Cells(nFila, 6) = cRefinac
          .Range(.Cells(nFila, 6), .Cells(nFila, 6)).HorizontalAlignment = xlCenter
          .Cells(nFila, 7) = nCuotasApr
          .Range(.Cells(nFila, 7), .Cells(nFila, 7)).HorizontalAlignment = xlRight
          .Cells(nFila, 8) = nAnosVig
          .Range(.Cells(nFila, 8), .Cells(nFila, 8)).HorizontalAlignment = xlRight
          .Cells(nFila, 9) = nMesVig
          .Range(.Cells(nFila, 9), .Cells(nFila, 9)).HorizontalAlignment = xlRight
          .Cells(nFila, 10) = nMontoCol
          .Range(.Cells(nFila, 10), .Cells(nFila, 10)).HorizontalAlignment = xlRight
          .Cells(nFila, 11) = nSaldoCap
          .Range(.Cells(nFila, 11), .Cells(nFila, 11)).HorizontalAlignment = xlRight
          .Cells(nFila, 12) = nPorcentaje
          .Range(.Cells(nFila, 12), .Cells(nFila, 12)).HorizontalAlignment = xlRight
          .Cells(nFila, 16) = nDiasMaximo
          .Range(.Cells(nFila, 16), .Cells(nFila, 16)).HorizontalAlignment = xlRight
          '.Cells(nFila, 13) = nDiasAtraso
          .Range(.Cells(nFila, 13), .Cells(nFila, 13)).HorizontalAlignment = xlRight
          SetearValoresNegativos nDiasAtraso, nFila, 13
          '.Cells(nFila, 14) = nDiasAtraso1
          .Range(.Cells(nFila, 14), .Cells(nFila, 14)).HorizontalAlignment = xlRight
          SetearValoresNegativos nDiasAtraso1, nFila, 14
          .Cells(nFila, 15) = ncCalifAnt
          .Range(.Cells(nFila, 15), .Cells(nFila, 15)).HorizontalAlignment = xlRight
          .Cells(nFila, 16) = nCalifVig
          .Range(.Cells(nFila, 16), .Cells(nFila, 16)).HorizontalAlignment = xlRight
          .Cells(nFila, 17) = nDiasMaximo
          .Range(.Cells(nFila, 17), .Cells(nFila, 17)).HorizontalAlignment = xlRight
          .Cells(nFila, 18) = nCalif
          .Range(.Cells(nFila, 18), .Cells(nFila, 18)).HorizontalAlignment = xlRight
          .Cells(nFila, 19) = nMoneda
          .Range(.Cells(nFila, 19), .Cells(nFila, 19)).HorizontalAlignment = xlRight
          .Cells(nFila, 20) = nProducto
          .Range(.Cells(nFila, 20), .Cells(nFila, 20)).HorizontalAlignment = xlRight
          .Cells(nFila, 21) = nAgencia
          .Range(.Cells(nFila, 21), .Cells(nFila, 21)).HorizontalAlignment = xlRight
           nFila = nFila + 1
            rs.MoveNext
        Loop
        Set rs = Nothing
        xlHoja1.PageSetup.Zoom = 75
        RaiseEvent Progress(9, 10)
        xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
        xlLibro.Close
        xlAplicacion.Quit
        RaiseEvent Progress(10, 10)
        MsgBox "Se genero el archivo " & App.path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"
        Set xlAplicacion = Nothing
    End With
    Exit Sub
ErrHandler:
    If Not objConecta Is Nothing Then Set objConecta = Nothing
    If Not fs Is Nothing Then Set fs = Nothing
    MsgBox "Se ha producido un error en el sistema", vbInformation, "AVISO"
End Sub

Public Sub Reporte_ClienteRFA()
    Dim rs As ADODB.Recordset
    Dim fs As New Scripting.FileSystemObject
    Dim nFila As Integer
    Dim nCol As Integer
    Dim lbExisteHoja As Boolean
    Dim lsCadena As String
    Dim objConecta As DConecta
    Dim ldFechaFinMes As Date
    Dim lsArchivo As String
    Dim lbExcel As Boolean
    Dim lNomHoja As String
    Dim sSQL As String
    
    Dim sNombre As String
    Dim sCodCliente As String
    Dim sCodCredito As String
    Dim dFecVig As Date
    Dim nCapDesembol As Double
    Dim nSalCap As Double
    Dim nPorcPago As Double
    Dim nCuotas As Integer
    Dim nPorCuotas  As Double
    Dim dFecPago As Date
    Dim cNroRef As String
    Dim nDiasAtraso As Integer
    Dim nDiasAtraso1 As Integer
    Dim nCalifAnt As Integer
    Dim nCalifAct As Integer
    Dim nCalifTemp As Integer
    Dim nCalifRef As Integer
    Dim nCalif As Integer
    Dim nMoneda As Integer
    Dim nProducto As Integer
    Dim nCodAgencia As Integer
    
    On Error GoTo ErrHandler
    
    ldFechaFinMes = GetUltimaFechaCierre
    Set objConecta = New DConecta
    RaiseEvent ShowProgress
    
    lsNomHoja = "Client.RFA " & "-" & CStr(ldFechaFinMes)
    RaiseEvent Progress(3, 10)
    
    lsArchivo = "CredRfa" & Format(ldFechaFinMes, "ddmmyyyy") & ".XLS"
    Set fs = New Scripting.FileSystemObject
    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    
    lbExisteHoja = False
    For Each xlHoja1 In xlLibro.Worksheets
        If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            xlHoja1.Range("A1", "AZ10000") = ""
            lbExisteHoja = True
            Exit For
        End If
        
    Next
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    
    nFila = 2
    xlHoja1.PageSetup.Zoom = 80
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlAplicacion.Range("A1:R1000").Font.Size = 9
    
    With xlHoja1
        .Range("A1").ColumnWidth = 30
        .Range("B1").ColumnWidth = 20
        .Range("C1").ColumnWidth = 20
        .Range("D1").ColumnWidth = 15
        .Range("E1").ColumnWidth = 20
        .Range("F1").ColumnWidth = 20
        .Range("G1").ColumnWidth = 20
        .Range("H1").ColumnWidth = 20
        .Range("I1").ColumnWidth = 20
        .Range("J1").ColumnWidth = 20
        .Range("K1").ColumnWidth = 10
        .Range("L1").ColumnWidth = 20
        .Range("M1").ColumnWidth = 20
        .Range("N1").ColumnWidth = 15
        .Range("O1").ColumnWidth = 15
        .Range("P1").ColumnWidth = 15
        .Range("Q1").ColumnWidth = 15
        .Range("R1").ColumnWidth = 15
        .Range("S1").ColumnWidth = 15
        .Range("T1").ColumnWidth = 20
        .Range("U1").ColumnWidth = 15
        .Range("V1").ColumnWidth = 15
        
        .Cells(nFila, 1) = gsNomCmac
        .Cells(nFila, 4) = "Creditos RFA " & CStr(ldFechaFinMes)
        .Cells(nFila, 7) = ""
        .Range(.Cells(nFila, 1), .Cells(nFila, 21)).Font.Bold = True
        
        nFila = nFila + 2
        .Cells(nFila, 1) = "Nombre del Cliente"
        .Cells(nFila, 2) = "Cód. Cliente"
        .Cells(nFila, 3) = "Cód. Crédito"
        .Cells(nFila, 4) = "Fec. Vigencia"
        .Cells(nFila, 5) = "Cap. Desem"
        .Cells(nFila, 6) = "Saldo Cap."
        .Cells(nFila, 7) = "% Pagado"
        .Cells(nFila, 8) = "Cuotas Apr."
        .Cells(nFila, 9) = "% Cuotas"
        .Cells(nFila, 10) = "Fec. Ult. Pag."
        .Cells(nFila, 11) = "Nor./Ref."
        .Cells(nFila, 12) = "Dias Atr. (CV)"
        .Cells(nFila, 13) = "Días Atr. (CA)"
        .Cells(nFila, 14) = "Calif. Ant"
        .Cells(nFila, 15) = "Calif. Act."
        .Cells(nFila, 16) = "Calif. Tmp"
        .Cells(nFila, 17) = "Calif. RFA"
        .Cells(nFila, 18) = "Calif."
        .Cells(nFila, 19) = "Moneda"
        .Cells(nFila, 20) = "Producto"
        .Cells(nFila, 21) = "Cód. Agencia"
                     
        sSQL = "select d.cPersNombre, "
       sSQL = sslq & "d.cPersCod,"
       sSQL = sSQL & "c.CtaCod, "
       sSQL = sSQL & "e.dVigencia, "
       sSQL = sSQL & "e.nMontoCol, "
       sSQL = sSQL & "a.nSaldoCap, "
       sSQL = sSQL & "isnull(e.nCuotas,0) as nCuotasApr, "
       sSQL = sSQL & "'' as PorcCuotasApr ," ' porcentaje de cuotas aprobadas
       sSQL = sSQL & " FecUtlimaPago=(select top 1 dVenc from dbo.ColocCalendario where cctacod=e.cCtaCod and nColocCalendApl=1 order by dVenc desc), "
       sSQL = sSQL & "b.cRefinan, "
       sSQL = sSQL & "b.nDiasAtraso, "
       sSQL = sSQL & "f.nDiasAtraso as nDiasAtraso1, "
       sSQL = sSQL & "b.cCalGen, "
       sSQL = sslq & "f.cCalGen as cCalGen1, "
       sSQL = sSQL & "'' as nCalifTemp," '--CalifTemp
       sSQL = sSQL & "'' as nCalifRfa," '--CalifRfa
       sSQL = sSQL & "b.cCalEval, "
       sSQL = sSQL & "substring(a.cCtaCod,9,1) as nMoneda, "
       sSQL = sSQL & "substring(a.cCtaCod,4,3) as nProducto,"
       sSQL = sSQL & "substring(a.cCtaCod, 3, 2) As nAgencia "
       sSQL = sSQL & "from colocacrefinanc a "
       sSQL = sSQL & "inner join dbcmacicaconsol..ColocCalifProv b on a.cctacod=b.cctacod "
       sSQL = sslq & "inner join PersonaProducto c on c.cCtaCod=b.cCtaCod and nPrdPersRelac=20 " ' -- titular
       sSQL = sSQL & "inner join Persona d on d.cCtaCod=c.cCtaCod "
       sSQL = sSQL & "left join Colocaciones e on e.cCtaCod=d.cCtaCod "
       sSQL = sSQL & "inner join dbcmacicaconsol..ColocCalifProvTotal f on f.cCtaCod=e.cCtaCod and datepart(mm,f.dFecVig)=datepart(mm,getdate)-1 "
       sSQL = sSQL & " where a.bRFA=1  and  b.nprdestado in ('2020','2021','2022','2030','2031','2032','2201','2101','2104','2106','2107') "
       
                
       objConecta.AbreConexion
       Set rs = objConecta.CargaRecordSet(sSQL)
       objConecta.CierraConexion
       RaiseEvent Progress(6, 10)
       nFila = nFila + 1
       Do Until rs.EOF
        sNombre = IIf(IsNull(rs!cPersNombre), "", rs!cPersNombre)
        sCodCliente = IIf(IsNull(rs!cPersCod), "", rs!cPersCod)
        sCodCredito = IIf(IsNull(rs!cCtaCod), "", rs!cCtaCod)
        dFecVig = IIf(IsNull(rs!dVigencia), "", rs!dVigencia)
        nCapDesembol = IIf(IsNull(rs!nMontoCol), 0, rs!nMontoCol)
        nSaldoCap = IIf(IsNull(rs!nSaldoCap), 0, rs!nSaldoCap)
        nCuotas = IIf(IsNull(rs!nCuotasApr), 0, rs!nCuotasApr)
        nPorcPago = ((nCapDesembol - nSaldoCap) / nCapDesembol) * 100
        dFecPago = IIf(IsNull(rs!FecUtlimaPago), "", rs!FecUtlimaPago)
        cNroRef = IIf(IsNull(rs!cRefinan), "", rs!cRefinan)
        nDiasAtraso = IIf(IsNull(rs!nDiasAtraso), 0, rs!nDiasAtraso)
        nDiasAtraso1 = IIf(IsNull(rs!nDiasAtraso1), 0, rs!nDiasAtraso1)
        nCalifAct = IIf(IsNull(rs!cCalGen), 0, rs!cCalGen)
        nCalifAnt = IIf(IsNull(rs!cCalGen1), 0, rs!cCalGen1)
        nCalif = IIf(IsNull(rs!cCalEval), 0, rs!cCalEval)
        nMoneda = IIf(IsNull(rs!nMoneda), 0, rs!nMoneda)
        nProducto = IIf(IsNull(rs!nProducto), 0, rs!nProducto)
        nCodAgencia = IIf(IsNull(rs!nAgencia), 0, rs!nAgencia)
        
        
        .Cells(nFila, 1) = sNombre
        .Range(.Cells(nFila, 1), .Cells(nFila, 1)).HorizontalAlignment = xlCenter
        .Cells(nFila, 2) = sCodCliente
        .Range(.Cells(nFila, 2), .Cells(nFila, 2)).HorizontalAlignment = xlRight
        .Cells(nFila, 3) = sCodCredito
        .Range(.Cells(nFila, 3), .Cells(nFila, 3)).HorizontalAlignment = xlRight
        .Cells(nFila, 4) = dFecVig
        .Range(.Cells(nFila, 4), .Cells(nFila, 4)).HorizontalAlignment = xlCenter
        .Cells(nFila, 5) = nCapDesembol
        .Range(.Cells(nFila, 5), .Cells(nFila, 5)).HorizontalAlignment = xlRight
        .Cells(nFila, 6) = nSaldoCap
        .Range(.Cells(nFila, 6), .Cells(nFila, 6)).HorizontalAlignment = xlRight
        .Cells(nFila, 7) = nPorcPago
        .Range(.Cells(nFila, 7), .Cells(nFila, 7)).HorizontalAlignment = xlRight
        .Cells(nFila, 8) = nCuotas
        .Range(.Cells(nFila, 8), .Cells(nFila, 8)).HorizontalAlignment = xlRight
        .Cells(nFila, 9) = "" ' porcentaje de las cuotas
        .Range(.Cells(nFila, 9), .Cells(nFila, 9)).HorizontalAlignment = xlRight
        .Range(.Cells(nFila, 9), .Cells(nFila, 9)).ColumnWidth = 0
        .Cells(nFila, 10) = dFecPago
        .Range(.Cells(nFila, 10), .Cells(nFila, 10)).HorizontalAlignment = xlRight
        .Cells(nFila, 11) = cNroRef
        .Range(.Cells(nFila, 11), .Cells(nFila, 11)).HorizontalAlignment = xlCenter
        .Cells(nFila, 12) = nDiasAtraso
        .Range(.Cells(nFila, 12), .Cells(nFila, 12)).HorizontalAlignment = xlRight
        .Cells(nFila, 13) = nDiasAtraso1
        .Range(.Cells(nFila, 13), .Cells(nFila, 13)).HorizontalAlignment = xlRight
        .Cells(nFila, 14) = nCalifAnt
        .Range(.Cells(nFila, 14), .Cells(nFila, 14)).HorizontalAlignment = xlRight
        .Cells(nFila, 15) = nCalifAct
        .Range(.Cells(nFila, 15), .Cells(nFila, 15)).HorizontalAlignment = xlRight
        .Cells(nFila, 16) = ""
        .Range(.Cells(nFila, 16), .Cells(nFila, 16)).HorizontalAlignment = xlRight
        .Range(.Cells(nFila, 16), .Cells(nFila, 16)).ColumnWidth = 0
        .Cells(nFila, 17) = ""
        .Range(.Cells(nFila, 17), .Cells(nFila, 17)).HorizontalAlignment = xlRight
        .Range(.Cells(nFila, 17), .Cells(nFila, 17)).ColumnWidth = 0
        .Cells(nFila, 18) = nCalif
        .Range(.Cells(nFila, 18), .Cells(nFila, 18)).HorizontalAlignment = xlRight
        .Cells(nFila, 19) = nMoneda
        .Range(.Cells(nFila, 19), .Cells(nFila, 19)).HorizontalAlignment = xlRight
        .Cells(nFila, 20) = nProducto
        .Range(.Cells(nFila, 20), .Cells(nFila, 20)).HorizontalAlignment = xlRight
        .Cells(nFila, 21) = nCodAgencia
        .Range(.Cells(nFila, 21), .Cells(nFila, 21)).HorizontalAlignment = xlRight
        rs.MoveNext
       Loop
    End With
    
    xlHoja1.PageSetup.Zoom = 75
    RaiseEvent ShowProgress
    xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
    'cierra el libro de trabajo
    xlLibro.Close
    ' cierre Microsoft Excel con el método QUIT
    xlAplicacion.Quit
    'libero los objetos
    RaiseEvent Progress(10, 10)
    MsgBox "Se genero el archivo en " & App.path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing
    lbExcel = True
    RaiseEvent CloseProgress
    Exit Sub
ErrHandler:
    
    If Not fs Is Nothing Then Set fs = Nothing
    If objConecta Is Nothing Then Set objConecta = Nothing
    MsgBox "Se ha producido un error en el sistema", vbCritical, "AVISO"
End Sub

Sub SetearValoresNegativos(ByVal nValor As Double, ByVal nFilaInicio As Integer, nIndex As Integer)
    With xlHoja1
        If nValor < 0 Then
            .Cells(nFilaInicio, nIndex) = "(" & (-1) * nValor & ")"
            .Range(.Cells(nFilaInicio, nIndex), .Cells(nFilaInicio, nIndex)).Font.Color = &H80&
        End If
    End With
End Sub
