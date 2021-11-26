VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMColRecRConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim coConex As COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling
Private Sub Class_Initialize()
    Dim loIni As COMConecta.DCOMClasIni
    Dim csConexion As String
    Dim csNegocio As String
    Dim csCentralPer As String
    Dim csCentralCom As String
    Dim csCentralImg As String
    Dim csAdminist As String

    Set loIni = New COMConecta.DCOMClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing

Set coConex = New COMConecta.DCOMConecta
If coConex.AbreConexion(csConexion) = False Then
    Call oError.RaiseError(oError.MyUnhandledError, "DColRecRConsulta:Initialize. Error en Conexion a Base de datos")
End If
End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

Public Function dObtieneDatosCabeceraRecuperacion(ByVal psctacod As String) As ADODB.Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lssql As String

Set lrs = New ADODB.Recordset

On Error GoTo dError

'MAVM 20100610 BAS II
'lssql = " SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion AS cEstado, Coloc.cUltimaActualizacion, T1.cConsDescripcion AS cRelacion, UPPER(T2.cConsDescripcion) " & _
'        " AS cProducto, UPPER(T3.cConsDescripcion) AS cMoneda, PP.cPersCod, PERS.cPersNombre, COLOC.dVigencia AS Fecha, PersID.cPersIDnro, " & _
'        " CR.cMetLiquid, cTipo=case when CR.nTipCj=1 then 'JUDICIAL' ELSE 'EXTRAJUDICIAL' END, " & _
'        " sCondicion=case when P.nPrdEstado in('2202', '2204','2206') then 'CASTIGADO' ELSE 'JUDICIAL' END, " & _
'        " sSEstado=case when P.nPrdEstado in('2201', '2202','2205','2206') then 'VIGENTE' ELSE 'CANCELADO' END " & _
'        " FROM ProductoPersona PP INNER JOIN Producto P ON P.cCtaCod = PP.cCtaCod " & _
'        " INNER JOIN Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN " & _
'        " Constante T2 ON SUBSTRING(PP.cCtaCod, 6, 3) = CONVERT(Varchar(3), T2.nConsValor) INNER JOIN " & _
'        " Constante T3 ON SUBSTRING(PP.cCtaCod, 9, 1) = CONVERT(Varchar(1), T3.nConsValor) INNER JOIN " & _
'        " Persona PERS ON PP.cPersCod = PERS.cPersCod INNER JOIN Colocaciones COLOC ON P.cCtaCod = COLOC.cCtaCod INNER JOIN " & _
'        " ColocRecup CR ON COLOC.cCtaCod = CR.cCtaCod LEFT OUTER JOIN PersID ON PERS.cPersCod = PersID.cPersCod " & _
'        " WHERE (PP.cCtaCod='" & psctacod & "') AND (T1.nConsCod = 3002) AND (T.nConsCod = 3001) " & _
'        " AND (T2.nConsCod = 1001) AND (T3.nConsCod = 1011) " & _
'        " AND (P.cCtaCod LIKE '_____%') AND (P.nPrdEstado IN (2201, 2202, 2203, 2204, 2205,2206)) AND (PP.nPrdPersRelac = 20)  "

'JUEZ 20150304 Convertido a SP **********************************************
'lssql = " SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion AS cEstado, Coloc.cUltimaActualizacion, T1.cConsDescripcion AS cRelacion, isnull(CN4.nConsValor,0) nTipoProdCod, ISNULL(CN4.cConsDescripcion,'') as cTipoProdDescrip " & _
'        " , UPPER(T3.cConsDescripcion) AS cMoneda, PP.cPersCod, PERS.cPersNombre, COLOC.dVigencia AS Fecha, PersID.cPersIDnro, " & _
'        " CR.cMetLiquid, cTipo=case when CR.nTipCj=1 then 'JUDICIAL' ELSE 'EXTRAJUDICIAL' END, " & _
'        " sCondicion=case when P.nPrdEstado in('2202', '2204','2206') then 'CASTIGADO' ELSE 'JUDICIAL' END, " & _
'        " sSEstado=case when P.nPrdEstado in('2201', '2202','2205','2206') then 'VIGENTE' ELSE 'CANCELADO' END " & _
'        " FROM ProductoPersona PP INNER JOIN Producto P ON P.cCtaCod = PP.cCtaCod " & _
'        " INNER JOIN Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN " & _
'        " Constante T3 ON SUBSTRING(PP.cCtaCod, 9, 1) = CONVERT(Varchar(1), T3.nConsValor) INNER JOIN " & _
'        " Persona PERS ON PP.cPersCod = PERS.cPersCod INNER JOIN Colocaciones COLOC ON P.cCtaCod = COLOC.cCtaCod INNER JOIN " & _
'        " Constante CN4 ON convert(int,substring(cTpoProdCod,1,2)+'0') = CN4.nConsValor AND CN4.nConsCod = " & gTpoProducto & " INNER JOIN " & _
'        " ColocRecup CR ON COLOC.cCtaCod = CR.cCtaCod LEFT OUTER JOIN PersID ON PERS.cPersCod = PersID.cPersCod " & _
'        " WHERE (PP.cCtaCod='" & psctacod & "') AND (T1.nConsCod = 3002) AND (T.nConsCod = 3001) " & _
'        " AND (T3.nConsCod = 1011) " & _
'        " AND (P.cCtaCod LIKE '_____%') AND (P.nPrdEstado IN (2201, 2202, 2203, 2204, 2205,2206)) AND (PP.nPrdPersRelac = 20)  "
        
lssql = "EXEC stp_sel_ObtieneDatosCabeceraRecuperacion '" & psctacod & "'"
'END JUEZ *******************************************************************

Set lrs = coConex.CargaRecordSet(lssql)
 
Set dObtieneDatosCabeceraRecuperacion = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosCabeceraRecuperacion>>", Err.Description
    
End Function

Public Function dObtieneDatosTotalesRecuperacion(ByVal psctacod As String) As ADODB.Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lssql As String

Set lrs = New ADODB.Recordset

On Error GoTo dError

'avmm 12-06-2006
lssql = "select sum(K.CapitalActual) as CapitalActual, sum(K.InteresActual) as InteresActual, sum(K.MoraActual) as MoraActual, " & _
        " sum(K.GastoActual) as GastoActual, sum(K.CapitalPagado) as CapitalPagado, sum(K.InteresPagado) as InteresPagado, sum(K.MoraPagado) as MoraPAgado, " & _
        " sum(k.GastoPAgado) As GastoPAgado From ( " & _
        "   SELECT Producto.nSaldo AS CapitalActual, ColocRecup.nSaldoIntComp AS InteresActual, " & _
        "   ColocRecup.nSaldoIntMor AS MoraActual, ColocRecup.nSaldoGasto AS GastoActual, " & _
        "   0 as CapitalPagado, 0 as InteresPagado, 0 as MoraPagado, 0 as GastoPagado " & _
        "   FROM ColocRecup INNER JOIN Producto ON ColocRecup.cCtaCod = Producto.cCtaCod " & _
        "   WHERE (ColocRecup.cCtaCod='" & psctacod & "') " & _
        " Union All " & _
        "   Select 0 as CapitalActual, 0 as InteresActual, 0 As MoraActual, 0 as GastoActual, " & _
        "   IsNull((SELECT SUM(nmonto) " & _
        "           FROM MovColDet MCD " & _
        "           INNER JOIN MOV M on MCD.nMovNro =M.nMovNro and nMovFlag = 0"
 lssql = lssql & _
        "           WHERE MCD.cCtaCod = P.cCtaCod AND " & _
        "           MCD.nPrdConceptoCod = '3000' and MCD.cOpecod like '130[234]%' ), 0) AS CapitalPagado,  " & _
        " IsNull ((SELECT SUM(nmonto)" & _
        "          FROM movcoldet MCD " & _
        "          INNER JOIN MOV M on MCD.nMovNro =M.nMovNro and nMovFlag = 0" & _
        "          WHERE MCD.cCtaCod = P.cCtaCod AND" & _
        "          MCD.nPrdConceptoCod = '3100' and MCD.cOpecod like '130[234]%'), 0) AS InteresPagado, " & _
        " IsNull ((SELECT SUM(nmonto) " & _
        "          FROM MovColDet MCD   " & _
        "          INNER JOIN MOV M on MCD.nMovNro =M.nMovNro and nMovFlag = 0 " & _
        "          WHERE MCD.cCtaCod = P.cCtaCod  AND" & _
        "          MCD.nPrdConceptoCod = '3101' and MCD.cOpecod like '130[234]%' ), 0) AS MoraPagado, " & _
        " IsNull ((SELECT SUM(nmonto)" & _
        "          FROM movcoldet MCD" & _
        "          INNER JOIN MOV M on MCD.nMovNro =M.nMovNro and nMovFlag = 0 " & _
        "          WHERE MCD.cCtaCod = P.cCtaCod AND " & _
        "          MCD.nPrdConceptoCod like '32%' and MCD.cOpecod like '130[234]%' ), 0) AS GastoPagado " & _
        " FROM Producto P where p.cctacod='" & psctacod & "') K "
        
Set lrs = coConex.CargaRecordSet(lssql)
 
Set dObtieneDatosTotalesRecuperacion = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosTotalesRecuperacion>>", Err.Description
    
End Function

Public Function dObtieneGastosRecuperacion(ByVal psctacod As String) As ADODB.Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lssql As String

Set lrs = New ADODB.Recordset

On Error GoTo dError

lssql = " SELECT '' as Item, convert(varchar(20),ColocRecupGastos.dAsigna,103) fecha, " & _
        " ProductoConcepto.cDescripcion, ColocRecupGastos.nMonto AS Importe, " & _
        " ColocRecupGastos.cMotivoGasto AS Origen FROM ColocRecupGastos INNER JOIN " & _
        " ProductoConcepto ON ColocRecupGastos.nPrdConceptoCod = ProductoConcepto.nPrdConceptoCod " & _
        " WHERE (ColocRecupGastos.cCtaCod='" & psctacod & "') AND ColocRecupGastos.nColocRecGastoEstado not in ( " & gColRecGastoEstEliminado & " ) " & _
        " order by  ColocRecupGastos.dAsigna "
        
Set lrs = coConex.CargaRecordSet(lssql)
 
Set dObtieneGastosRecuperacion = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneGastosRecuperacion>>", Err.Description
    
End Function

Public Function dObtieneTotalesGastosRecuperacion(ByVal psctacod As String) As ADODB.Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lssql As String

Set lrs = New ADODB.Recordset

On Error GoTo dError

lssql = " SELECT  Isnull(SUM(nMonto),0) AS Importe, Isnull(SUM(nMontoPagado),0) AS Pagado " & _
        " From dbo.ColocRecupGastos  " & _
        " Where cCtaCod ='" & psctacod & "' AND nColocRecGastoEstado not in ( " & gColRecGastoEstEliminado & " ) "
        
Set lrs = coConex.CargaRecordSet(lssql)
 
Set dObtieneTotalesGastosRecuperacion = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosContrato>>", Err.Description
    
End Function

Public Function dObtieneListaAmortizaciones(ByVal psctacod As String) As ADODB.Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lssql As String

Set lrs = New ADODB.Recordset

On Error GoTo dError

lssql = "SELECT DISTINCT '' AS Item,  convert(varchar(10),convert(datetime,substring(cMovNro,1,8)),103) as fecha, OpeTpo.cOpeDesc + space(25) + Mov.cOpeCod  AS Operacion, '' AS Importe, ISNULL " & _
        " ((SELECT Sum(nmonto) FROM movcoldet MCD WHERE Mov.nMovNro = MCD.nMovNro And MCD.cCtaCod = MDet.cCtaCod AND MCD.nPrdConceptoCod in('3000','1401','1402','1403','1404','1405','1406','1407','1408' )  AND MCD.cOpeCod = MDet.cOpeCod), 0) AS Capital, " & _
        "  ISNULL " & _
        " ((SELECT Sum(nmonto) FROM movcoldet MCD WHERE Mov.nMovNro = MCD.nMovNro And MCD.cCtaCod = MDet.cCtaCod AND MCD.nPrdConceptoCod = '3100' AND MCD.cOpeCod = MDet.cOpeCod), 0) AS Interes, ISNULL " & _
        " ((SELECT Sum(nmonto) FROM movcoldet MCD WHERE Mov.nMovNro = MCD.nMovNro And MCD.cCtaCod = MDet.cCtaCod AND MCD.nPrdConceptoCod = '3101' AND MCD.cOpeCod = MDet.cOpeCod), 0) AS Mora, ISNULL " & _
        " ((SELECT Sum(nmonto) FROM movcoldet MCD WHERE Mov.nMovNro = MCD.nMovNro And MCD.cCtaCod = MDet.cCtaCod AND MCD.nPrdConceptoCod like '32%' AND MCD.cOpeCod = MDet.cOpeCod), 0) AS Gastos, " & _
        " '' AS Saldo, Mov.cMovNro as Numero " & _
        " FROM MovColDet MDet INNER JOIN OpeTpo ON MDet.cOpeCod = OpeTpo.cOpeCod INNER JOIN " & _
        " MovCol ON MDet.nMovNro = MovCol.nMovNro AND MDet.cOpeCod = MovCol.cOpeCod AND MDet.cCtaCod = MovCol.cCtaCod AND " & _
        " MDet.nNroCalen = MovCol.nNroCalen INNER JOIN Mov ON MovCol.nMovNro = Mov.nMovNro " & _
        " WHERE (MDet.cCtaCod='" & psctacod & "') AND MDet.cOpeCod like '13%' and nMovFlag = 0 ORDER BY Mov.cMovNro "

Set lrs = coConex.CargaRecordSet(lssql)
 

Set dObtieneListaAmortizaciones = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneListaAmortizaciones>>", Err.Description
    
End Function

Public Function dObtieneDatosExpediente(ByVal psctacod As String) As ADODB.Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lssql As String

Set lrs = New ADODB.Recordset

On Error GoTo dError

lssql = "SELECT  distinct ColocRecupExpediente.cNumExp, PERSO.cPersNombre,PERSO.cPersCod, " & _
        " PERSO.cPersDireccDomicilio , " & _
        " sMoneda=case when ColocRecupExpediente.nMoneda=1 then 'SOLES' else 'DOLARES' end, " & _
        " sTipoComision=case when ColocRecupComision.nTipComis=1 then 'P' else 'M' end, ColocRecupComision.nValor, " & _
        " ColocRecupExpediente.nMonPetit, " & _
        " (select KONS.cconsdescripcion from constante KONS where KONS.nconscod='3303' " & _
        "     and KONS.nconsvalor=ColocRecupExpediente.nViaProce) as sViaProcesal, " & _
        " (select KONS.cconsdescripcion from constante KONS where KONS.nconscod='3308' " & _
        "     and KONS.nconsvalor=ColocRecupExpediente.nEstadoProceso) as sEstadoProceso " & _
        " FROM ColocRecupExpediente INNER JOIN " & _
        " ColocRecup ON ColocRecupExpediente.cCtaCod = ColocRecup.cCtaCod INNER JOIN " & _
        " Colocaciones ON ColocRecup.cCtaCod = Colocaciones.cCtaCod INNER JOIN " & _
        " Producto ON Colocaciones.cCtaCod = Producto.cCtaCod INNER JOIN " & _
        " ProductoPersona PP ON Producto.cCtaCod = PP.cCtaCod INNER JOIN " & _
        " ColocRecupComision ON ColocRecup.nComisionCod = ColocRecupComision.nComisionCod " & _
        " inner join productopersona PP1 on PP1.cctacod=PP.cctacod inner join persona perso on perso.cperscod=pp1.cperscod " & _
        " WHERE (ColocRecupExpediente.cCtaCod='" & psctacod & "') and  pp1.NPRDPERSRELAC = 30 "
          
Set lrs = coConex.CargaRecordSet(lssql)
 
Set dObtieneDatosExpediente = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosExpediente>>", Err.Description

End Function


Public Function dObtieneDatosGenerales(ByVal psctacod As String) As ADODB.Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lssql As String

Set lrs = New ADODB.Recordset

On Error GoTo dError

lssql = "SELECT DISTINCT ProductoPersona.nPrdPersRelac, Colocaciones.dVigencia AS FechaPrestamo, Colocaciones.nMontoCol, " & _
        " ColocRecup.dIngRecup AS FechaIngreso, p.cPersNombre AS Analista, ColocLineaCredito.cDescripcion,  " & _
        " nTasaInt = (SELECT ISNULL(nTasaInteres, 0) From ProductoTasaInteres LCT  WHERE LCT.cCtaCod =Colocaciones.cCtaCod and LCT.nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & " ) , " & _
        " nTasaIntMor = (SELECT ISNULL(nTasaInteres, 0) From ProductoTasaInteres LCT  WHERE LCT.cCtaCod = Colocaciones.cCtaCod and LCT.nPrdTasaInteres = " & gColocLineaCredTasasIntMoratNormal & " ), Colocaciones.cUltimaActualizacion" & _
        " FROM Colocaciones INNER JOIN ProductoPersona ON Colocaciones.cCtaCod = ProductoPersona.cCtaCod " & _
        " INNER JOIN ColocRecup ON Colocaciones.cCtaCod = ColocRecup.cCtaCod  " & _
        " INNER JOIN Persona p ON ProductoPersona.cPersCod = p.cPersCod  " & _
        " INNER JOIN ColocLineaCredito ON Colocaciones.cLineaCred = ColocLineaCredito.cLineaCred " & _
        " WHERE (ProductoPersona.nPrdPersRelac = 28) AND (Colocaciones.cCtaCod = '" & psctacod & "')"
 
Set lrs = coConex.CargaRecordSet(lssql)
 
Set dObtieneDatosGenerales = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosGenerales>>", Err.Description

End Function

'***TODOCOMPLETA  VEIFICAR FUNCION
Function GetExcelVersion() As String
'    Dim sVersion As String
'    Dim objExceLServer As Object
'
'
'    sVersion = "0"
'
'    Set objExceLServer = GetObject("Excel.Application.10")
'    If Err.Number = 0 Then
'        sVersion = "10.0"
'        Exit Function
'    End If
'
'    Err.Clear
'
'    Set objExceLServer = CreateObject("Excel.Application.9")
'    If Err.Number = 0 Then
'        sVersion = "9.0"
'        Exit Function
'    End If
'    Err.Clear
'
'    Set objExceLServer = CreateObject("Excel.Application.8")
'        If Err.Number = 0 Then
'            sVersion = "8.0"
'            Exit Function
'        End If
'    Err.Clear
'    GetExcelVersion = sVersion
End Function

Public Function Recup_PagosGestores(ByVal psFileName, ByVal pdFechaInicio As Date, ByVal pdFechaFin As Date) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim ssql As String
'    Dim m_Excel As Excel.Application
'    Dim oLibroExcel As Excel.Workbook
'    Dim oHojaExcel As Excel.Worksheet
    Dim sCuentas As String
    Dim sFechaInicio As String
    Dim sFechaFin As String
    Dim nCantidad As Integer
    Dim nCont As Integer
    
    Dim i As Integer
    
    sFechaInicio = Format(pdFechaInicio, "YYYYMMDD")
    sFechaFin = Format(pdFechaFin, "YYYYmmDD")
    
'    Set m_Excel = New Excel.Application
'    Set oLibroExcel = m_Excel.Workbooks.Open(psFileName)
'    Set oHojaExcel = oLibroExcel.Worksheets(1)
    
'    nCantidad = oHojaExcel.Cells(1, 1)
'    nCont = 0
'    For I = 2 To nCantidad
'        If oHojaExcel.Cells(I, 1) <> "" Then
'             If nCont = 0 Then
'                sCuentas = "'" & oHojaExcel.Cells(I, 1) & "'"
'             Else
'                sCuentas = sCuentas & ",'" & oHojaExcel.Cells(I, 1) & "'"
'             End If
'        End If
'        nCont = nCont + 1
'    Next I
'    oLibroExcel.Close
'    m_Excel.Application.Quit
            
    
    'configurando el excel de la consulta
'    Set cn = New ADODB.Connection
'    'cn.Provider = "Microsoft.Jet.OLEDB.4.0"
'    'cn.ConnectionString = " Data Source=" & psFileName & ";Extended Properties=Excel 9.0;"
'    cn.Open " provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & psFileName & ";Extended Properties=Excel 11.0;"
'    cn.Open
'    Set rs = cn.Execute("Select * From [Sheet$A1:A65000] ")
'    cn.Close
'    Set cn = Nothing
'
'    i = 0
'    Do Until rs.EOF
'        If i = 0 Then
'           If rs.Fields(0) <> "" Then
'                sCuentas = "'" & rs.Fields(0) & "'"
'           End If
'        Else
'            If rs.Fields(0) <> "" Then
'                sCuentas = ",'" & rs.Fields(0) & "'"
'            End If
'        End If
'        i = i + 1
'        rs.MoveNext
'    Loop
'    Set rs = Nothing
    
    ssql = "Select nMovNro,cCtaCod,cPersNombre,Capital,Interes,Mora,Gastos,Moneda,nSaldo,FechaPago "
    ssql = ssql & " From ("
    ssql = ssql & " Select X.nMovNro,"
    ssql = ssql & " X.cPersNombre,"
    ssql = ssql & " P.cCtaCod,"
    ssql = ssql & " Capital=isnull((Select Sum(nMonto) From MovColDet MD"
    ssql = ssql & " Where X.nMovNro=MD.nMovNro and  MD.cCtaCod=X.cCtaCod and MD. nPrdConceptoCod in (3000,1000,1109) and MD.cOpeCod not like '107[123456789]%'),0),"
    ssql = ssql & " Interes=isnull((Select Sum(nMonto) From MovColDet MD"
    ssql = ssql & " Where X.nMovNro=MD.nMovNro and  MD.cCtaCod=X.cCtaCod and MD.nPrdConceptoCod in (3100,1100,1107) and MD.cOpeCod not like '107[123456789]%'),0),"
    ssql = ssql & " Mora=isnull((Select Sum(nMonto) From MovColDet MD"
    ssql = ssql & " Where X.nMovNro=MD.nMovNro and  MD.cCtaCod=X.cCtaCod and MD.nPrdConceptoCod in (3101,1101,1108) and MD.cOpeCod not like '107[123456789]%'),0),"
    ssql = ssql & " Gastos=isnull((Select Sum(nMonto) From MovColDet MD"
    ssql = ssql & " Where X.nMovNro=MD.nMovNro and  MD.cCtaCod=X.cCtaCod and MD.nPrdConceptoCod Like '12%' and MD.cOpeCod not like '107[123456789]%'),0),"
    ssql = ssql & " Moneda=Case When SubString(X.cCtaCod,9,1)='1' Then 'Soles' Else 'Dolares' End,"
    ssql = ssql & " P.nSaldo,"
    ssql = ssql & " FechaPago=(Select  SubString(left(M.cMovNro,8),7,2)+'/'+SubString(left(M.cMovNro,8),5,2)+'/'+SubString(left(M.cMovNro,8),1,4) From Mov M  Where M.nMovNro=X.nMovNro)"
    ssql = ssql & " From (Select M.nMovNro,Pers.cPersNombre,MD.cCtaCod"
    ssql = ssql & " From Mov M"
    ssql = ssql & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
    ssql = ssql & " Inner Join ProductoPersona PP on PP.cCtaCod=MD.cCtaCod and PP.nPrdPersRelac=20"
    ssql = ssql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    ssql = ssql & " Where (M.nmovflag=0 or M.nMovFlag=5)and Left(m.cmovnro,8) between '" & sFechaInicio & "' and '" & sFechaFin & "' and"
    ssql = ssql & " MD.cOpeCod not like '107[123456789]%' and MD.cOpeCod<>'107002' and MD.nPrdConceptoCod not like '14%'"
    ssql = ssql & " Group By M.nMovNro,Pers.cPersNombre,MD.cCtaCod) X"
    ssql = ssql & " Inner Join producto p on P.cCtaCod=X.cCtaCod"
    'sSQL = sSQL & " Inner Join TablaGestores TG on TG.cCtaCod=P.cCtaCod"
    ssql = ssql & " Where P.cCtaCod in (" & sCuentas & ")) X"
    ssql = ssql & " Group BY nMovNro,cCtaCod,cPersNombre,Capital,Interes,Mora,Gastos,Moneda,nSaldo,FechaPago"
    ssql = ssql & " Order By FechaPago"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set Recup_PagosGestores = oConec.CargaRecordSet(ssql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaCreditosJudCast(ByVal psCodAgen As String, ByVal sTipoFiltro As String) As ADODB.Recordset
    Dim oConecta As COMConecta.DCOMConecta
    Dim ssql As String
    Dim sNPrdEstado As String
    
    If sTipoFiltro = "J" Then
        sNPrdEstado = "2201,2205"
    Else
        sNPrdEstado = "2202,2206"
    End If
    
    ssql = "Select P.cCtaCod,"
    ssql = ssql & " Cn.cConsDescripcion as cEstado,"
    ssql = ssql & " Pers.cPersNombre as cTitular,"
    ssql = ssql & " Pers.cPersDireccDomicilio as cDireccion,"
    ssql = ssql & " Pers1.cPersNombre as cGarante,"
    ssql = ssql & " Pers1.cPersDireccDomicilio as cDireccionGarante,"
    ssql = ssql & " Convert(Varchar(20),C.dVigencia,103) as dVigencia,"
    ssql = ssql & " C.nMontoCol,"
    ssql = ssql & " P.nSaldo,"
    ssql = ssql & " Cr.nSaldoIntComp,"
    ssql = ssql & " CR.nSaldoIntMor,"
    ssql = ssql & " CR.nSaldoGasto,"
    ssql = ssql & " TotalDeuda= P.nSaldo+ Cr.nSaldoIntComp+CR.nSaldoIntMor+CR.nSaldoGasto,"
    ssql = ssql & " Cn.cConsDescripcion as cViaProcesal,"
    ssql = ssql & " CRE.cNumExp,"
    ssql = ssql & " CR.dIngRecup,"
    ssql = ssql & " Pers2.cPersNombre as cJuzgado,"
    ssql = ssql & " Pers2.cPersNombre as cAbogado2"
    ssql = ssql & " From Producto P"
    ssql = ssql & " Inner Join Constante Cn on Cn.nConsCod=3001 and Cn.nConsValor=P.nPrdEstado and Cn.nConsValor<>3001"
    ssql = ssql & " Inner Join ProductoPersona PP on PP.cCtaCod=P.cCtaCod and PP.nPrdPersRelac=20"
    ssql = ssql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    ssql = ssql & " Left  Join ProductoPersona PP1 on PP1.cCtaCod=P.cCtaCod and PP1.nPrdPersRelac=25"
    ssql = ssql & " Left Join  Persona Pers1 on Pers1.cPersCod=PP1.cPersCod"
    ssql = ssql & " Inner Join Colocaciones C on C.cCtaCod=P.cCtaCod"
    ssql = ssql & " Inner Join ColocRecup CR on CR.cCtaCod=P.cCtaCod"
    ssql = ssql & " Left Join ColocRecupExpediente CRE on CRE.cCtacod=P.cCtaCod"
    ssql = ssql & " Left Join Constante Cn1 on Cn1.nConsCod=3003 and Cn1.nConsValor=CRE.nViaProce"
    ssql = ssql & " Left Join ProductoPersona PP2 on PP2.cCtaCod=P.cCtaCod and PP2.nPrdPersRelac=30"
    ssql = ssql & " Left Join Persona Pers2 on Pers2.cPerscod=PP2.cPersCod"
    ssql = ssql & " Left Join ProductoPersona PP3 on PP3.cCtaCod=P.cCtaCod and PP3.nPrdPersRelac=31"
    ssql = ssql & " Left Join Persona Pers3 on Pers3.cPerscod=PP3.cPersCod"
    ssql = ssql & " Where P.nPrdEstado in (" & sNPrdEstado & ") and SubString(P.cCtaCod,4,2)=" & psCodAgen

    Set oConecta = New COMConecta.DCOMConecta
    oConecta.AbreConexion
    Set ListaCreditosJudCast = oConecta.CargaRecordSet(ssql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function NumeroCreditosVencXAnalista(ByVal psCodAgen As String, ByVal pdFecha As Date) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim ssql As String
    Dim sFecha As String
    
    sFecha = Format(pdFecha, "MM/dd/yyyy")
    
    ssql = "Select Count(CS.cCtaCod) as nCantidad,RH.cUser"
    ssql = ssql & " From ColocacSaldo CS"
    ssql = ssql & " Inner Join ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
    ssql = ssql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    ssql = ssql & " Where SubString(CS.cCtaCod,4,2)='" & psCodAgen & "' and CS.nDiasAtraso=31 and CS.dFecha='" & sFecha & "' and"
    ssql = ssql & " CS.nPrdEstado in  (2020,2021,2022,2030,2031,2032)"
    ssql = ssql & " Group By RH.cUser"
    ssql = ssql & " Order By RH.cUser"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set NumeroCreditosVencXAnalista = oConec.CargaRecordSet(ssql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ListaAnalistaXAgenciaXCredito(ByVal psCodAgen As String, ByVal pdFechaInicial As Date, _
ByVal pdFechaFinal As Date, Optional ByVal pbCantidad As Boolean = False) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim ssql As String
    Dim sFechaInicial As String
    Dim sFechaFinal As String
    
    sFechaInicial = Format(pdFechaInicial, "MM/dd/yyyy")
    sFechaFinal = Format(pdFechaFinal, "MM/dd/yyyy")
    
    If pbCantidad = False Then
        ssql = "Select  RH.cUser"
        ssql = ssql & " From ColocacSaldo CS"
        ssql = ssql & " Inner JOin ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
        ssql = ssql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
        ssql = ssql & " Where CS.dFecha between '" & sFechaInicial & "' and '" & sFechaFinal & "' and SubString(CS.cCtaCod,4,2)='" & psCodAgen & "' and"
        ssql = ssql & " CS.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
        ssql = ssql & " Group By RH.cUser"
        ssql = ssql & " Order By RH.cUser"
   Else
        ssql = "Select  Count(Distinct RH.cUser) as nCantidad"
        ssql = ssql & " From ColocacSaldo CS"
        ssql = ssql & " Inner JOin ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
        ssql = ssql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
        ssql = ssql & " Where CS.dFecha between '" & sFechaInicial & "' and '" & sFechaFinal & "' and SubString(CS.cCtaCod,4,2)='" & psCodAgen & "' and"
        ssql = ssql & " CS.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
   End If
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaAnalistaXAgenciaXCredito = oConec.CargaRecordSet(ssql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaMontosVencidosXAnalista(ByVal psCodAgen As String, ByVal pdFecha As Date, _
ByVal pnTipoCambio As Double) As ADODB.Recordset
    Dim ssql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sFecha As String
    
    sFecha = Format(pdFecha, "MM/dd/yyyy")
    
    ssql = "Select Sum(Case When SubString(CS.cCtaCod,9,1)='1' Then CS.nSaldoCap Else CS.nSaldoCap*" & pnTipoCambio & " End) as nCantidad,RH.cUser"
    ssql = ssql & " From ColocacSaldo CS"
    ssql = ssql & " Inner Join ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
    ssql = ssql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    ssql = ssql & " Where SubString(CS.cCtaCod,4,2)='" & psCodAgen & "' and CS.nDiasAtraso=31 and CS.dFecha='" & sFecha & "' and"
    ssql = ssql & " CS.nPrdEstado in  (2020,2021,2022,2030,2031,2032)"
    ssql = ssql & " Group By RH.cUser"
    ssql = ssql & " Order By RH.cUser"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaMontosVencidosXAnalista = oConec.CargaRecordSet(ssql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerSaldoVencidoFecha(ByVal pdFecha As Date, ByVal pnTipoCambio As Double, _
                                          ByVal psCodAgen As String, ByVal psCodUser As String) As ADODB.Recordset

 Dim ssql As String
 Dim oConec As COMConecta.DCOMConecta
 Dim sFecha As String
 
 sFecha = Format(pdFecha, "MM/dd/yyyy")
 
 ssql = "Select Sum(Case When SubString(CS.cCtaCod,9,1)=1 Then CS.nSaldoCap Else CS.nSaldoCap*3.261 End ) as nMonto"
 ssql = ssql & " From ColocacSaldo CS"
 ssql = ssql & " Inner Join ColocacCred C on C.cCtaCod=CS.cCtaCod"
 ssql = ssql & " Inner Join ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
 ssql = ssql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
 ssql = ssql & " Where RH.cUser='" & psCodUser & "' and CS.dFecha='" & sFecha & "' and CS.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
 ssql = ssql & " and C.nDiasAtraso>=31 and SubString(CS.cCtaCod,4,2)='" & psCodAgen & "'"

 Set oConec = New COMConecta.DCOMConecta
 oConec.AbreConexion
 Set ObtenerSaldoVencidoFecha = oConec.CargaRecordSet(ssql)
 oConec.CierraConexion
 Set oConec = Nothing
End Function

Public Function ObtenerMontoInicialesVencidos(ByVal pdFecha As Date, ByVal pnTipoCambio As Double, _
ByVal psCodAgen As String, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim ssql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sFecha As String
    
    sFecha = Format(pdFecha, "MM/dd/yyyy")
    
    ssql = "Select  Sum(Case When SubString(CS.cCtaCod,9,1)='1' Then CS.nSaldoCap Else  CS.nSaldoCap*" & pnTipoCambio & " End) as nMonto,"
    ssql = ssql & " RH.cUser"
    ssql = ssql & " From ColocacSaldo CS"
    ssql = ssql & " Inner Join ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
    ssql = ssql & " Inner Join ColocacCred C on C.cCtaCod=PP.cCtaCod"
    ssql = ssql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    ssql = ssql & " Where CS.dFecha='" & sFecha & "' and CS.nDiasAtraso>=31 and (C.cRFA ='NOR' Or C.cRFA Is Null) and"
    ssql = ssql & " SubString(CS.cCtaCod,4,2)='" & psCodAgen & "' and SubString(CS.cCtacod,9,1)='" & pnMoneda & "'"
    ssql = ssql & "  And CS.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
    ssql = ssql & " Group By RH.cUser"
    ssql = ssql & " Order BY  RH.cUser"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerMontoInicialesVencidos = oConec.CargaRecordSet(ssql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ObtenerMontoIngresadoXAnalista(ByVal pdFecha As Date, ByVal pnTipoCambio As Double, _
ByVal psCodAgen As String, ByVal psCodUser As String, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim ssql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sFecha As String
    
    sFecha = Format(pdFecha, "MM/dd/yyyy")
    
    ssql = "Select isnull(Sum(Case When SubString(CS.cCtaCod,9,1)='1'Then CS.nSaldoCap Else CS.nSaldoCap*3.315 End),0) as nMonto"
    ssql = ssql & " From ColocacSaldo CS"
    ssql = ssql & " Inner Join ColocacCred C on C.cCtacod=CS.cCtaCod"
    ssql = ssql & " Inner Join ProductoPersona PP on PP.cCtacod=C.cCtaCod and PP.nPrdPersRelac=28"
    ssql = ssql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    ssql = ssql & " Where dFecha='" & sFecha & "' and RH.cUser='" & psCodUser & "' and (C.cRFA='NOR' Or C.cRFA is NULL) and"
    ssql = ssql & " SubString(CS.cCtaCod,4,2)='" & psCodAgen & "' and SubString(CS.cCtaCod,9,1)=" & pnMoneda
    ssql = ssql & " and CS.nDiasAtraso=31 and CS.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerMontoIngresadoXAnalista = oConec.CargaRecordSet(ssql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerMontoPagadoXAnalista(ByVal psCodAgen As String, ByVal pnMoneda As Integer, _
ByVal psCodUser As String, ByVal pdFechaInicial As Date, ByVal pdFechaFinal As Date, _
ByVal pnTipoCambio As Double) As ADODB.Recordset
    Dim ssql As String
    Dim oConec As COMConecta.DCOMConecta
    Dim sFechaInicial As String
    Dim sFechaFinal As String
    
    sFechaInicial = Format(pdFechaInicial, "YYYYMMDD")
    sFechaFinal = Format(pdFechaFinal, "YYYYMMDD")
    
    ssql = "Select Isnull(Sum(Case When SubString(MD.cCtaCod,9,1)='1' Then MD.nMonto Else MD.nMonto*" & pnTipoCambio & " End),0) as nMonto"
    ssql = ssql & " From Mov M"
    ssql = ssql & " Inner Join MovcolDet MD on MD.nMovNro=M.nMovNro"
    ssql = ssql & " Inner Join ProductoPersona PP on PP.cCtaCod=MD.cCtaCod and PP.nPrdPersRelac=28"
    ssql = ssql & " Inner Join RRHH  RH on RH.cPersCod=PP.cPersCod"
    ssql = ssql & " Where SubString(MD.cCtaCod,4,2)='" & psCodAgen & "' and SubString(MD.cCtaCod,9,1)='" & pnMoneda & "' and RH.cUser='" & psCodUser & "' and"
    ssql = ssql & " (M.nMovFlag=0 or M.nMovFlag=5) and MD.nPrdConceptoCod=1000 and"
    ssql = ssql & " MD.cOpeCod not in ('100101','100102','100103','100104','100105') and MD.cOpeCod not like '107[123456789]%' and"
    ssql = ssql & " (Left(M.cMovNro,8) Between '" & sFechaInicial & "' and '" & sFechaFinal & "')"
       
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerMontoPagadoXAnalista = oConec.CargaRecordSet(ssql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ObtenerMontosTotalesXAnalista(ByVal pdFecha As Date, ByVal psCodAgen As String, _
ByVal psMoneda As String, ByVal pnTipoCambio As Double, ByVal psCodUser As String) As ADODB.Recordset
    Dim oConec As COMConecta.DCOMConecta
    Dim ssql As String
    Dim sFechaInicial As String
    
    sFechaInicial = Format(pdFecha, "MM/dd/YYYY")
    
    ssql = "Select Sum(Case When SubString(CS.cCtaCod,9,1)='1' Then CS.nSaldoCap Else CS.nSaldoCap*" & pnTipoCambio & " End ) as nMonto"
    ssql = ssql & " From ColocacSaldo CS"
    ssql = ssql & " Inner Join ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
    ssql = ssql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    ssql = ssql & " Inner Join ColocacCred C on C.cCtaCod=CS.cCtaCod"
    ssql = ssql & " Where SubString(CS.cCtacod,4,2)='" & psCodAgen & "' and CS.dFecha='" & sFechaInicial & "'  and"
    ssql = ssql & " (C.cRFA='NOR' or C.cRFA Is Null) and RH.cUser='" & psCodUser & "'"

    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerMontosTotalesXAnalista = oConec.CargaRecordSet(ssql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerMontosFinalesXAnalista(ByVal pdFecha As Date, ByVal psCodAgen As String, _
ByVal psMoneda As String, ByVal pnTipoCambio As Double, ByVal psCodUser As String) As ADODB.Recordset

    Dim oConec As COMConecta.DCOMConecta
    Dim ssql As String
    Dim sFechaInicial As String
    
    sFechaInicial = Format(pdFecha, "MM/dd/yyyy")
    
    ssql = "Select Sum(Case When SubString(CS.cCtaCod,9,1)='1' Then CS.nSaldoCap Else CS.nSaldoCap End ) as nMonto"
    ssql = ssql & " From ColocacSaldo CS"
    ssql = ssql & " Inner Join ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
    ssql = ssql & " Inner Join RrHh RH on RH.cPersCod=PP.cPersCod"
    ssql = ssql & " Inner Join ColocacCred C on C.cCtaCod=CS.cCtaCod"
    ssql = ssql & " Where SubString(CS.cCtacod,4,2)='" & psCodAgen & "' and CS.dFecha='" & sFechaInicial & "' and"
    ssql = ssql & " (C.cRFA='NOR' or C.cRFA Is Null) and RH.cUser='" & psCodUser & "'  and CS.nDiasAtraso>=31 and"
    ssql = ssql & " CS.nPrdEstado in (2020,2021,2022,2030,2031,2032) and SubString(CS.cCtaCod,9,1)='" & psMoneda & "'"
    
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ObtenerMontosFinalesXAnalista = oConec.CargaRecordSet(ssql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ListaCreditosXAnalistaXAgencia(ByVal psCodAgen As String, ByVal psCodAnalista As String, _
ByVal psCodMoneda As Integer, Optional ByVal psTipoC As String = "", Optional ByVal pdFechaI As Date, _
Optional ByVal pdFechaF As Date) As ADODB.Recordset
    Dim ssql As String
    Dim oConec As COMConecta.DCOMConecta
    
    Dim oCons As COMDConstSistema.NCOMConstSistema
    Dim dFecSisTmp As Date

    Set oCons = New COMDConstSistema.NCOMConstSistema
    dFecSisTmp = CDate(oCons.LeeConstSistema(gConstSistFechaSistema))
    Set oCons = Nothing
    
    '****** BRGO BASILEA II  20100615
    
    ssql = "Select P.cCtaCod,"
    ssql = ssql & " Pers.cPersNombre as cTitular,"
    ssql = ssql & " Pers.cPersTelefono as cTitularTelefono1,"
    ssql = ssql & " Pers.cPersTelefono2 as cTitularTelefono2,"
    ssql = ssql & " Pers.cPersDireccDomicilio as cTitularDomicilio,"
    ssql = ssql & " cTitularDirecNegocio=(Select PFI.cRazSocDirecc"
    ssql = ssql & "                         From ColocFteIngreso CF"
    ssql = ssql & "                                 Inner Join PersFteIngreso PFI on CF.cNumFuente=PFI.cNumFuente"
    ssql = ssql & "                           Where CF.cCtaCod=P.cCtaCod),"
    ssql = ssql & " Pers1.cPersNombre as cNombreAval,"
    ssql = ssql & " Pers1.cPersDireccDomicilio as cAvalDomicilio,"
    ssql = ssql & " Pers1.cPersTelefono as cAvalTelefono1,"
    ssql = ssql & " Pers1.cPersTelefono2 as cAvalTelefono2,"
    ssql = ssql & " P.nSaldo as nSK,"
    ssql = ssql & " CR.nSaldoIntComp + dbo.ColocRecupCalculoInteresActual(P.cCtaCod,'" & Format(dFecSisTmp, "mm/dd/yyyy") & "') as nMontoComp,"
    ssql = ssql & " CR.nSaldoIntMor + dbo.ColocRecupCalculoMoraActual(P.cCtaCod,'" & Format(dFecSisTmp, "mm/dd/yyyy") & "') as nMontoMor,"
    ssql = ssql & " CR.nSaldoGasto as nGasto,"
    ssql = ssql & " FechaPago=(Select SubString(cMovNro,7,2)+'/'+SubString(cMovNro,5,2)+'/'+SubString(cMovNro,1,4)"
    ssql = ssql & "             From Mov Where nMovNro=(Select Max(nMovnro) From MovCol Where cCtacod=P.cCtaCod AND cOpeCod<>'131000')),"
    ssql = ssql & " MontoPago=(Select Sum(nMonto)"
    ssql = ssql & "            From MovCol Where nMovNro=(Select Max(nMovNro) From MovCol Where cCtaCod=P.cCtaCod AND cOpeCod<>'131000')),"
    ssql = ssql & " C.cAgeCodAct as cAgencia,"
    ssql = ssql & " SubString (P.cCtaCod,9,1) as cMoneda,"
    ssql = ssql & " PP2.cPersCod  as cAnalista"
    ssql = ssql & " From Producto P"
    ssql = ssql & " Inner Join Colocaciones C on C.cCtaCod = P.cCtaCod "
    ssql = ssql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=20"
    ssql = ssql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    ssql = ssql & " Left Join ProductoPersona PP1 on PP1.cCtacod=P.cCtaCod and PP1.nPrdPersRelac in (21,22,25)"
    ssql = ssql & " Left Join Persona Pers1 on Pers1.cPersCod=PP1.cPersCod"
    ssql = ssql & " Inner Join ColocRecup  CR on CR.cCtaCod=P.cCtaCod"
    ssql = ssql & " Inner Join ProductoPersona PP2 on PP2.cCtaCod=P.cCtaCod and PP2.nPrdPersRelac=28"
    ssql = ssql & " Where P.nPrdEstado in (2201,2205) and C.cAgeCodAct In (" & psCodAgen & ") and PP2.cPersCod in (" & psCodAnalista & ") and SubString(P.cCtaCod,9,1)= " & psCodMoneda
    ssql = ssql & " and ( dIngRecup >='" & Format(pdFechaI, "yyyymmdd") & "' and  dIngRecup <='" & Format(pdFechaF, "yyyymmdd") & "')"
    
    If Trim(psTipoC) <> "" Then
        ssql = ssql & " and nTipCj in " & "(" & psTipoC & ")"
    End If
    
    ssql = ssql & " Order By PP2.cPersCod, C.cAgeCodAct, SubString(P.cCtaCod,9,1),P.cCtaCod"
  
    '**********************
    Set oConec = New COMConecta.DCOMConecta
    oConec.AbreConexion
    Set ListaCreditosXAnalistaXAgencia = oConec.CargaRecordSet(ssql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

'JUEZ 20121013 *******************************************************************
Public Function dObtieneGastoRecuperacionXNroReg(ByVal psctacod As String, ByVal pnNroGastoCta As Integer) As ADODB.Recordset
Dim lrs As ADODB.Recordset
Dim lssql As String

Set lrs = New ADODB.Recordset
On Error GoTo dError

lssql = "SELECT convert(varchar(20),CRG.dAsigna,103) dFechaAsigna, "
lssql = lssql & " CRG.nPrdConceptoCod , PC.cDescripcion, CRG.nMonto, CRG.cMotivoGasto "
lssql = lssql & "FROM ColocRecupGastos CRG "
lssql = lssql & " INNER JOIN ProductoConcepto PC ON CRG.nPrdConceptoCod = PC.nPrdConceptoCod "
lssql = lssql & "WHERE (CRG.cCtaCod='" & psctacod & "') AND CRG.nNroGastoCta=" & pnNroGastoCta
        
Set lrs = coConex.CargaRecordSet(lssql)
Set dObtieneGastoRecuperacionXNroReg = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneGastoRecuperacionXNroReg>>", Err.Description
End Function

Public Function dObtieneDatosAsigCredRecup(ByVal psctacod As String) As ADODB.Recordset
Dim lrs As ADODB.Recordset
Dim lssql As String
Set lrs = New ADODB.Recordset
On Error GoTo dError

lssql = " SELECT P.cCtaCod, PP.cPersCod, P.nPrdEstado, C.nMontoCol, P.nSaldo, C.cUltimaActualizacion,  " _
    & " CRec.nSaldoIntComp, CRec.nSaldoIntMor, CRec.nSaldoGasto, CRec.nIntCompGen,CRec.nIntMoraGen, CRec.cNroNeg, " _
    & " CRec.nComisionCod, CRec.cMetLiquid, CRec.nTipCj, CRec.nDemanda, CRec.nNroCalen,  " _
    & " Pers.cPersCod, Pers.cPersNombre, " _
    & " nUltGas = (Select Isnull(max(nNroGastoCta),0) From ColocRecupGastos CRG Where CRG.cCtaCod = p.cCtaCod ) , " _
    & " nTasaInt = (SELECT ISNULL(nTasaInteres, 0) From ProductoTasaInteres LCT  WHERE LCT.cCtaCod = P.cCtaCod and LCT.nPrdTasaInteres = " & gColocLineaCredTasasIntCompNormal & " ) , " _
    & " nTasaIntMor = (SELECT ISNULL(nTasaInteres, 0) From ProductoTasaInteres LCT  WHERE LCT.cCtaCod = P.cCtaCod and LCT.nPrdTasaInteres = " & gColocLineaCredTasasIntMoratNormal & " ) " _
    & " FROM Producto P Inner Join Colocaciones C ON P.cCtaCod = C.cCtaCod " _
    & " Inner Join ColocRecup CRec ON C.cctacod = CRec.cctacod " _
    & " Inner Join ProductoPersona PP ON C.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = " & gColRelPersTitular _
    & " Inner Join Persona Pers ON Pers.cPersCod = PP.cPersCod " _
    & " WHERE P.cCtaCod ='" & psctacod & "'"

Set lrs = coConex.CargaRecordSet(lssql)
    
Set dObtieneDatosAsigCredRecup = lrs
Set lrs = Nothing
Exit Function
dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosPagoCredRecup>>", Err.Description
    
End Function
'END JUEZ ************************************************************************
