VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DColRecRConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String

Dim coConex As DConecta
Private Sub Class_Initialize()
    Dim loIni As COMConecta.DCOMClasIni
    
    Set loIni = New COMConecta.DCOMClasIni
        csConexion = loIni.CadenaConexion
        csNegocio = loIni.BaseNegocio
        csCentralPer = loIni.BasePersonas
        csCentralCom = loIni.BaseComunes
        csCentralImg = loIni.BaseImagenes
        csAdminist = loIni.BaseAdministracion
    Set loIni = Nothing

Set coConex = New DConecta
If coConex.AbreConexion(csConexion) = False Then
    Call RaiseError(MyUnhandledError, "DColRecRConsulta:Initialize. Error en Conexion a Base de datos")
End If
End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

Public Function dObtieneDatosCabeceraRecuperacion(ByVal psctacod As String) As Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = " SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion AS cEstado, Coloc.cUltimaActualizacion, T1.cConsDescripcion AS cRelacion, UPPER(T2.cConsDescripcion) " & _
        " AS cProducto, UPPER(T3.cConsDescripcion) AS cMoneda, PP.cPersCod, PERS.cPersNombre, COLOC.dVigencia AS Fecha, PersID.cPersIDnro, " & _
        " CR.cMetLiquid, cTipo=case when CR.nTipCj=1 then 'JUDICIAL' ELSE 'EXTRAJUDICIAL' END, " & _
        " sCondicion=case when P.nPrdEstado in('2202', '2204','2206') then 'CASTIGADO' ELSE 'JUDICIAL' END, " & _
        " sSEstado=case when P.nPrdEstado in('2201', '2202','2205','2206') then 'VIGENTE' ELSE 'CANCELADO' END " & _
        " FROM ProductoPersona PP INNER JOIN Producto P ON P.cCtaCod = PP.cCtaCod " & _
        " INNER JOIN Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN " & _
        " Constante T2 ON SUBSTRING(PP.cCtaCod, 6, 3) = CONVERT(Varchar(3), T2.nConsValor) INNER JOIN " & _
        " Constante T3 ON SUBSTRING(PP.cCtaCod, 9, 1) = CONVERT(Varchar(1), T3.nConsValor) INNER JOIN " & _
        " Persona PERS ON PP.cPersCod = PERS.cPersCod INNER JOIN Colocaciones COLOC ON P.cCtaCod = COLOC.cCtaCod INNER JOIN " & _
        " ColocRecup CR ON COLOC.cCtaCod = CR.cCtaCod LEFT OUTER JOIN PersID ON PERS.cPersCod = PersID.cPersCod " & _
        " WHERE (PP.cCtaCod='" & psctacod & "') AND (T1.nConsCod = 3002) AND (T.nConsCod = 3001) " & _
        " AND (T2.nConsCod = 1001) AND (T3.nConsCod = 1011) " & _
        " AND (P.cCtaCod LIKE '_____%') AND (P.nPrdEstado IN (2201, 2202, 2203, 2204, 2205,2206)) AND (PP.nPrdPersRelac = 20)  "

Set lrs = coConex.CargaRecordSet(lsSQL)
 
Set dObtieneDatosCabeceraRecuperacion = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosCabeceraRecuperacion>>", Err.Description
    
End Function

Public Function dObtieneDatosTotalesRecuperacion(ByVal psctacod As String) As Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = "select sum(K.CapitalActual) as CapitalActual, sum(K.InteresActual) as InteresActual, sum(K.MoraActual) as MoraActual, " & _
        " sum(K.GastoActual) as GastoActual, sum(K.CapitalPagado) as CapitalPagado, sum(K.InteresPagado) as InteresPagado, sum(K.MoraPagado) as MoraPAgado, " & _
        " sum(k.GastoPAgado) As GastoPAgado From ( " & _
        "   SELECT Producto.nSaldo AS CapitalActual, ColocRecup.nSaldoIntComp AS InteresActual, " & _
        "   ColocRecup.nSaldoIntMor AS MoraActual, ColocRecup.nSaldoGasto AS GastoActual, " & _
        "   0 as CapitalPagado, 0 as InteresPagado, 0 as MoraPagado, 0 as GastoPagado " & _
        "   FROM ColocRecup INNER JOIN Producto ON ColocRecup.cCtaCod = Producto.cCtaCod " & _
        "   WHERE (ColocRecup.cCtaCod='" & psctacod & "') " & _
        " Union All " & _
        "   Select 0 as CapitalActual, 0 as InteresActual, 0 As MoraActual, 0 as GastoActual, " & _
        "   IsNull((SELECT SUM(nmonto) FROM MovColDet MCD WHERE MCD.cCtaCod = P.cCtaCod AND " & _
        "           MCD.nPrdConceptoCod = '3000' and MCD.cOpecod like '130[234]%' ), 0) AS CapitalPagado,  " & _
        " IsNull ((SELECT SUM(nmonto) FROM movcoldet MCD WHERE MCD.cCtaCod = P.cCtaCod AND " & _
        "          MCD.nPrdConceptoCod = '3100' and MCD.cOpecod like '130[234]%'), 0) AS InteresPagado, " & _
        " IsNull ((SELECT SUM(nmonto) FROM MovColDet MCD WHERE MCD.cCtaCod = P.cCtaCod  AND " & _
        "          MCD.nPrdConceptoCod = '3101' and MCD.cOpecod like '130[234]%' ), 0) AS MoraPagado, " & _
        " IsNull ((SELECT SUM(nmonto) FROM movcoldet MCD WHERE MCD.cCtaCod = P.cCtaCod AND " & _
        "          MCD.nPrdConceptoCod like '32%' and MCD.cOpecod like '130[234]%' ), 0) AS GastoPagado " & _
        " FROM Producto P where p.cctacod='" & psctacod & "') K "
        
Set lrs = coConex.CargaRecordSet(lsSQL)
 
Set dObtieneDatosTotalesRecuperacion = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosTotalesRecuperacion>>", Err.Description
    
End Function

Public Function dObtieneGastosRecuperacion(ByVal psctacod As String) As Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = " SELECT '' as Item, convert(varchar(20),ColocRecupGastos.dAsigna,103) fecha, " & _
        " ProductoConcepto.cDescripcion, ColocRecupGastos.nMonto AS Importe, " & _
        " ColocRecupGastos.cMotivoGasto AS Origen FROM ColocRecupGastos INNER JOIN " & _
        " ProductoConcepto ON ColocRecupGastos.nPrdConceptoCod = ProductoConcepto.nPrdConceptoCod " & _
        " WHERE (ColocRecupGastos.cCtaCod='" & psctacod & "') AND ColocRecupGastos.nColocRecGastoEstado not in ( " & gColRecGastoEstEliminado & " ) " & _
        " order by  ColocRecupGastos.dAsigna "
        
Set lrs = coConex.CargaRecordSet(lsSQL)
 
Set dObtieneGastosRecuperacion = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneGastosRecuperacion>>", Err.Description
    
End Function

Public Function dObtieneTotalesGastosRecuperacion(ByVal psctacod As String) As Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = " SELECT  Isnull(SUM(nMonto),0) AS Importe, Isnull(SUM(nMontoPagado),0) AS Pagado " & _
        " From dbo.ColocRecupGastos  " & _
        " Where cCtaCod ='" & psctacod & "' AND nColocRecGastoEstado not in ( " & gColRecGastoEstEliminado & " ) "
        
Set lrs = coConex.CargaRecordSet(lsSQL)
 
Set dObtieneTotalesGastosRecuperacion = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosContrato>>", Err.Description
    
End Function


Public Function dObtieneListaAmortizaciones(ByVal psctacod As String) As Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = "SELECT DISTINCT '' AS Item,  convert(varchar(10),convert(datetime,substring(cMovNro,1,8)),103) as fecha, OpeTpo.cOpeDesc + space(25) + Mov.cOpeCod  AS Operacion, '' AS Importe, ISNULL " & _
        " ((SELECT Sum(nmonto) FROM movcoldet MCD WHERE Mov.nMovNro = MCD.nMovNro And MCD.cCtaCod = MDet.cCtaCod AND MCD.nPrdConceptoCod in('3000','1401','1402','1403','1404','1405','1406','1407','1408' )  AND MCD.cOpeCod = MDet.cOpeCod), 0) AS Capital, " & _
        "  ISNULL " & _
        " ((SELECT Sum(nmonto) FROM movcoldet MCD WHERE Mov.nMovNro = MCD.nMovNro And MCD.cCtaCod = MDet.cCtaCod AND MCD.nPrdConceptoCod = '3100' AND MCD.cOpeCod = MDet.cOpeCod), 0) AS Interes, ISNULL " & _
        " ((SELECT Sum(nmonto) FROM movcoldet MCD WHERE Mov.nMovNro = MCD.nMovNro And MCD.cCtaCod = MDet.cCtaCod AND MCD.nPrdConceptoCod = '3101' AND MCD.cOpeCod = MDet.cOpeCod), 0) AS Mora, ISNULL " & _
        " ((SELECT Sum(nmonto) FROM movcoldet MCD WHERE Mov.nMovNro = MCD.nMovNro And MCD.cCtaCod = MDet.cCtaCod AND MCD.nPrdConceptoCod like '32%' AND MCD.cOpeCod = MDet.cOpeCod), 0) AS Gastos, " & _
        " '' AS Saldo, Mov.cMovNro as Numero " & _
        " FROM MovColDet MDet INNER JOIN OpeTpo ON MDet.cOpeCod = OpeTpo.cOpeCod INNER JOIN " & _
        " MovCol ON MDet.nMovNro = MovCol.nMovNro AND MDet.cOpeCod = MovCol.cOpeCod AND MDet.cCtaCod = MovCol.cCtaCod AND " & _
        " MDet.nNroCalen = MovCol.nNroCalen INNER JOIN Mov ON MovCol.nMovNro = Mov.nMovNro " & _
        " WHERE (MDet.cCtaCod='" & psctacod & "') AND MDet.cOpeCod like '13%' ORDER BY Mov.cMovNro "

Set lrs = coConex.CargaRecordSet(lsSQL)
 

Set dObtieneListaAmortizaciones = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneListaAmortizaciones>>", Err.Description
    
End Function

Public Function dObtieneDatosExpediente(ByVal psctacod As String) As Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = "SELECT  distinct ColocRecupExpediente.cNumExp, PERSO.cPersNombre,PERSO.cPersCod, " & _
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
          
Set lrs = coConex.CargaRecordSet(lsSQL)
 
Set dObtieneDatosExpediente = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosExpediente>>", Err.Description

End Function


Public Function dObtieneDatosGenerales(ByVal psctacod As String) As Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = "SELECT DISTINCT ProductoPersona.nPrdPersRelac, Colocaciones.dVigencia AS FechaPrestamo, Colocaciones.nMontoCol, " & _
        " ColocRecup.dIngRecup AS FechaIngreso, p.cPersNombre AS Analista, ColocLineaCredito.cDescripcion,  " & _
        " nTasaIntComp = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " & _
        "                 WHERE LCT.cLineaCred = Colocaciones.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntCompNormal & " ),  " & _
        " nTasaIntMorat = (SELECT ISNULL(nTasaIni, 0) From ColocLineaCreditoTasa LCT " & _
        "                 WHERE LCT.cLineaCred = Colocaciones.cLineaCred and LCT.nColocLinCredTasaTpo = " & gColocLineaCredTasasIntMoratNormal & " )  " & _
        " FROM Colocaciones INNER JOIN ProductoPersona ON Colocaciones.cCtaCod = ProductoPersona.cCtaCod " & _
        " INNER JOIN ColocRecup ON Colocaciones.cCtaCod = ColocRecup.cCtaCod  " & _
        " INNER JOIN Persona p ON ProductoPersona.cPersCod = p.cPersCod  " & _
        " INNER JOIN ColocLineaCredito ON Colocaciones.cLineaCred = ColocLineaCredito.cLineaCred " & _
        " WHERE (ProductoPersona.nPrdPersRelac = 28) AND (Colocaciones.cCtaCod = '" & psctacod & "')"
 
Set lrs = coConex.CargaRecordSet(lsSQL)
 
Set dObtieneDatosGenerales = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosGenerales>>", Err.Description

End Function

Function GetExcelVersion() As String
    Dim sVersion As String
    Dim objExceLServer As Object
    
    
    sVersion = "0"
    
    Set objExceLServer = GetObject("Excel.Application.10")
    If Err.Number = 0 Then
        sVersion = "10.0"
        Exit Function
    End If
    
    Err.Clear
    
    Set objExceLServer = CreateObject("Excel.Application.9")
    If Err.Number = 0 Then
        sVersion = "9.0"
        Exit Function
    End If
    Err.Clear
    
    Set objExceLServer = CreateObject("Excel.Application.8")
        If Err.Number = 0 Then
            sVersion = "8.0"
            Exit Function
        End If
    Err.Clear
    GetExcelVersion = sVersion
End Function

Public Function Recup_PagosGestores(ByVal psFileName, ByVal pdFechaInicio As Date, ByVal pdFechaFin As Date) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    Dim m_Excel As Excel.Application
    Dim oLibroExcel As Excel.Workbook
    Dim oHojaExcel As Excel.Worksheet
    Dim sCuentas As String
    Dim sFechaInicio As String
    Dim sFechaFin As String
    Dim nCantidad As Integer
    Dim NCONT As Integer
    
    Dim i As Integer
    
    sFechaInicio = Format(pdFechaInicio, "YYYYMMDD")
    sFechaFin = Format(pdFechaFin, "YYYYmmDD")
    
    Set m_Excel = New Excel.Application
    Set oLibroExcel = m_Excel.Workbooks.Open(psFileName)
    Set oHojaExcel = oLibroExcel.Worksheets(1)
    
    nCantidad = oHojaExcel.Cells(1, 1)
    NCONT = 0
    For i = 2 To nCantidad
        If oHojaExcel.Cells(i, 1) <> "" Then
             If NCONT = 0 Then
                sCuentas = "'" & oHojaExcel.Cells(i, 1) & "'"
             Else
                sCuentas = sCuentas & ",'" & oHojaExcel.Cells(i, 1) & "'"
             End If
        End If
        NCONT = NCONT + 1
    Next i
    oLibroExcel.Close
    m_Excel.Application.Quit
            
    
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
    
    sSql = "Select nMovNro,cCtaCod,cPersNombre,Capital,Interes,Mora,Gastos,Moneda,nSaldo,FechaPago "
    sSql = sSql & " From ("
    sSql = sSql & " Select X.nMovNro,"
    sSql = sSql & " X.cPersNombre,"
    sSql = sSql & " P.cCtaCod,"
    sSql = sSql & " Capital=isnull((Select Sum(nMonto) From MovColDet MD"
    sSql = sSql & " Where X.nMovNro=MD.nMovNro and  MD.cCtaCod=X.cCtaCod and MD. nPrdConceptoCod in (3000,1000,1109) and MD.cOpeCod not like '107[123456789]%'),0),"
    sSql = sSql & " Interes=isnull((Select Sum(nMonto) From MovColDet MD"
    sSql = sSql & " Where X.nMovNro=MD.nMovNro and  MD.cCtaCod=X.cCtaCod and MD.nPrdConceptoCod in (3100,1100,1107) and MD.cOpeCod not like '107[123456789]%'),0),"
    sSql = sSql & " Mora=isnull((Select Sum(nMonto) From MovColDet MD"
    sSql = sSql & " Where X.nMovNro=MD.nMovNro and  MD.cCtaCod=X.cCtaCod and MD.nPrdConceptoCod in (3101,1101,1108) and MD.cOpeCod not like '107[123456789]%'),0),"
    sSql = sSql & " Gastos=isnull((Select Sum(nMonto) From MovColDet MD"
    sSql = sSql & " Where X.nMovNro=MD.nMovNro and  MD.cCtaCod=X.cCtaCod and MD.nPrdConceptoCod Like '12%' and MD.cOpeCod not like '107[123456789]%'),0),"
    sSql = sSql & " Moneda=Case When SubString(X.cCtaCod,9,1)='1' Then 'Soles' Else 'Dolares' End,"
    sSql = sSql & " P.nSaldo,"
    sSql = sSql & " FechaPago=(Select  SubString(left(M.cMovNro,8),7,2)+'/'+SubString(left(M.cMovNro,8),5,2)+'/'+SubString(left(M.cMovNro,8),1,4) From Mov M  Where M.nMovNro=X.nMovNro)"
    sSql = sSql & " From (Select M.nMovNro,Pers.cPersNombre,MD.cCtaCod"
    sSql = sSql & " From Mov M"
    sSql = sSql & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=MD.cCtaCod and PP.nPrdPersRelac=20"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Where (M.nmovflag=0 or M.nMovFlag=5)and Left(m.cmovnro,8) between '" & sFechaInicio & "' and '" & sFechaFin & "' and"
    sSql = sSql & " MD.cOpeCod not like '107[123456789]%' and MD.cOpeCod<>'107002' and MD.nPrdConceptoCod not like '14%'"
    sSql = sSql & " Group By M.nMovNro,Pers.cPersNombre,MD.cCtaCod) X"
    sSql = sSql & " Inner Join producto p on P.cCtaCod=X.cCtaCod"
    'sSQL = sSQL & " Inner Join TablaGestores TG on TG.cCtaCod=P.cCtaCod"
    sSql = sSql & " Where P.cCtaCod in (" & sCuentas & ")) X"
    sSql = sSql & " Group BY nMovNro,cCtaCod,cPersNombre,Capital,Interes,Mora,Gastos,Moneda,nSaldo,FechaPago"
    sSql = sSql & " Order By FechaPago"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set Recup_PagosGestores = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaCreditosJudCast(ByVal psCodAgen As String, ByVal sTipoFiltro As String) As Recordset
    Dim oConecta As DConecta
    Dim sSql As String
    Dim sNPrdEstado As String
    
    If sTipoFiltro = "J" Then
        sNPrdEstado = "2201,2205"
    Else
        sNPrdEstado = "2202,2206"
    End If
    
    sSql = "Select P.cCtaCod,"
    sSql = sSql & " Cn.cConsDescripcion as cEstado,"
    sSql = sSql & " Pers.cPersNombre as cTitular,"
    sSql = sSql & " Pers.cPersDireccDomicilio as cDireccion,"
    sSql = sSql & " Pers1.cPersNombre as cGarante,"
    sSql = sSql & " Pers1.cPersDireccDomicilio as cDireccionGarante,"
    sSql = sSql & " Convert(Varchar(20),C.dVigencia,103) as dVigencia,"
    sSql = sSql & " C.nMontoCol,"
    sSql = sSql & " P.nSaldo,"
    sSql = sSql & " Cr.nSaldoIntComp,"
    sSql = sSql & " CR.nSaldoIntMor,"
    sSql = sSql & " CR.nSaldoGasto,"
    sSql = sSql & " TotalDeuda= P.nSaldo+ Cr.nSaldoIntComp+CR.nSaldoIntMor+CR.nSaldoGasto,"
    sSql = sSql & " Cn.cConsDescripcion as cViaProcesal,"
    sSql = sSql & " CRE.cNumExp,"
    sSql = sSql & " CR.dIngRecup,"
    sSql = sSql & " Pers2.cPersNombre as cJuzgado,"
    sSql = sSql & " Pers2.cPersNombre as cAbogado2"
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join Constante Cn on Cn.nConsCod=3001 and Cn.nConsValor=P.nPrdEstado and Cn.nConsValor<>3001"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=P.cCtaCod and PP.nPrdPersRelac=20"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Left  Join ProductoPersona PP1 on PP1.cCtaCod=P.cCtaCod and PP1.nPrdPersRelac=25"
    sSql = sSql & " Left Join  Persona Pers1 on Pers1.cPersCod=PP1.cPersCod"
    sSql = sSql & " Inner Join Colocaciones C on C.cCtaCod=P.cCtaCod"
    sSql = sSql & " Inner Join ColocRecup CR on CR.cCtaCod=P.cCtaCod"
    sSql = sSql & " Left Join ColocRecupExpediente CRE on CRE.cCtacod=P.cCtaCod"
    sSql = sSql & " Left Join Constante Cn1 on Cn1.nConsCod=3003 and Cn1.nConsValor=CRE.nViaProce"
    sSql = sSql & " Left Join ProductoPersona PP2 on PP2.cCtaCod=P.cCtaCod and PP2.nPrdPersRelac=30"
    sSql = sSql & " Left Join Persona Pers2 on Pers2.cPerscod=PP2.cPersCod"
    sSql = sSql & " Left Join ProductoPersona PP3 on PP3.cCtaCod=P.cCtaCod and PP3.nPrdPersRelac=31"
    sSql = sSql & " Left Join Persona Pers3 on Pers3.cPerscod=PP3.cPersCod"
    sSql = sSql & " Where P.nPrdEstado in (" & sNPrdEstado & ") and SubString(P.cCtaCod,4,2)=" & psCodAgen

    Set oConecta = New DConecta
    oConecta.AbreConexion
    Set ListaCreditosJudCast = oConecta.CargaRecordSet(sSql)
    oConecta.CierraConexion
    Set oConecta = Nothing
End Function

Public Function NumeroCreditosVencXAnalista(ByVal psCodAgen As String, ByVal pdFecha As Date) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    Dim sFecha As String
    
    sFecha = Format(pdFecha, "MM/dd/yyyy")
    
    sSql = "Select Count(CS.cCtaCod) as nCantidad,RH.cUser"
    sSql = sSql & " From ColocacSaldo CS"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
    sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    sSql = sSql & " Where SubString(CS.cCtaCod,4,2)='" & psCodAgen & "' and CS.nDiasAtraso=31 and CS.dFecha='" & sFecha & "' and"
    sSql = sSql & " CS.nPrdEstado in  (2020,2021,2022,2030,2031,2032)"
    sSql = sSql & " Group By RH.cUser"
    sSql = sSql & " Order By RH.cUser"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set NumeroCreditosVencXAnalista = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ListaAnalistaXAgenciaXCredito(ByVal psCodAgen As String, ByVal pdFechaInicial As Date, _
ByVal pdFechaFinal As Date, Optional ByVal pbCantidad As Boolean = False) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    Dim sFechaInicial As String
    Dim sFechaFinal As String
    
    sFechaInicial = Format(pdFechaInicial, "MM/dd/yyyy")
    sFechaFinal = Format(pdFechaFinal, "MM/dd/yyyy")
    
    If pbCantidad = False Then
        sSql = "Select  RH.cUser"
        sSql = sSql & " From ColocacSaldo CS"
        sSql = sSql & " Inner JOin ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
        sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
        sSql = sSql & " Where CS.dFecha between '" & sFechaInicial & "' and '" & sFechaFinal & "' and SubString(CS.cCtaCod,4,2)='" & psCodAgen & "' and"
        sSql = sSql & " CS.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
        sSql = sSql & " Group By RH.cUser"
        sSql = sSql & " Order By RH.cUser"
   Else
        sSql = "Select  Count(Distinct RH.cUser) as nCantidad"
        sSql = sSql & " From ColocacSaldo CS"
        sSql = sSql & " Inner JOin ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
        sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
        sSql = sSql & " Where CS.dFecha between '" & sFechaInicial & "' and '" & sFechaFinal & "' and SubString(CS.cCtaCod,4,2)='" & psCodAgen & "' and"
        sSql = sSql & " CS.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
   End If
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaAnalistaXAgenciaXCredito = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ListaMontosVencidosXAnalista(ByVal psCodAgen As String, ByVal pdFecha As Date, _
ByVal pnTipoCambio As Double) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    Dim sFecha As String
    
    sFecha = Format(pdFecha, "MM/dd/yyyy")
    
    sSql = "Select Sum(Case When SubString(CS.cCtaCod,9,1)='1' Then CS.nSaldoCap Else CS.nSaldoCap*" & pnTipoCambio & " End) as nCantidad,RH.cUser"
    sSql = sSql & " From ColocacSaldo CS"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
    sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    sSql = sSql & " Where SubString(CS.cCtaCod,4,2)='" & psCodAgen & "' and CS.nDiasAtraso=31 and CS.dFecha='" & sFecha & "' and"
    sSql = sSql & " CS.nPrdEstado in  (2020,2021,2022,2030,2031,2032)"
    sSql = sSql & " Group By RH.cUser"
    sSql = sSql & " Order By RH.cUser"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaMontosVencidosXAnalista = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerSaldoVencidoFecha(ByVal pdFecha As Date, ByVal pnTipoCambio As Double, _
                                          ByVal psCodAgen As String, ByVal psCodUser As String) As Recordset

 Dim sSql As String
 Dim oConec As DConecta
 Dim sFecha As String
 
 sFecha = Format(pdFecha, "MM/dd/yyyy")
 
 sSql = "Select Sum(Case When SubString(CS.cCtaCod,9,1)=1 Then CS.nSaldoCap Else CS.nSaldoCap*3.261 End ) as nMonto"
 sSql = sSql & " From ColocacSaldo CS"
 sSql = sSql & " Inner Join ColocacCred C on C.cCtaCod=CS.cCtaCod"
 sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
 sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
 sSql = sSql & " Where RH.cUser='" & psCodUser & "' and CS.dFecha='" & sFecha & "' and CS.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
 sSql = sSql & " and C.nDiasAtraso>=31 and SubString(CS.cCtaCod,4,2)='" & psCodAgen & "'"

 Set oConec = New DConecta
 oConec.AbreConexion
 Set ObtenerSaldoVencidoFecha = oConec.CargaRecordSet(sSql)
 oConec.CierraConexion
 Set oConec = Nothing
End Function

Public Function ObtenerMontoInicialesVencidos(ByVal pdFecha As Date, ByVal pnTipoCambio As Double, _
ByVal psCodAgen As String, ByVal pnMoneda As Integer) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    Dim sFecha As String
    
    sFecha = Format(pdFecha, "MM/dd/yyyy")
    
    sSql = "Select  Sum(Case When SubString(CS.cCtaCod,9,1)='1' Then CS.nSaldoCap Else  CS.nSaldoCap*" & pnTipoCambio & " End) as nMonto,"
    sSql = sSql & " RH.cUser"
    sSql = sSql & " From ColocacSaldo CS"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
    sSql = sSql & " Inner Join ColocacCred C on C.cCtaCod=PP.cCtaCod"
    sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    sSql = sSql & " Where CS.dFecha='" & sFecha & "' and CS.nDiasAtraso>=31 and (C.cRFA ='NOR' Or C.cRFA Is Null) and"
    sSql = sSql & " SubString(CS.cCtaCod,4,2)='" & psCodAgen & "' and SubString(CS.cCtacod,9,1)='" & pnMoneda & "'"
    sSql = sSql & "  And CS.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
    sSql = sSql & " Group By RH.cUser"
    sSql = sSql & " Order BY  RH.cUser"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ObtenerMontoInicialesVencidos = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ObtenerMontoIngresadoXAnalista(ByVal pdFecha As Date, ByVal pnTipoCambio As Double, _
ByVal psCodAgen As String, ByVal psCodUser As String, ByVal pnMoneda As Integer) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    Dim sFecha As String
    
    sFecha = Format(pdFecha, "MM/dd/yyyy")
    
    sSql = "Select isnull(Sum(Case When SubString(CS.cCtaCod,9,1)='1'Then CS.nSaldoCap Else CS.nSaldoCap*3.315 End),0) as nMonto"
    sSql = sSql & " From ColocacSaldo CS"
    sSql = sSql & " Inner Join ColocacCred C on C.cCtacod=CS.cCtaCod"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtacod=C.cCtaCod and PP.nPrdPersRelac=28"
    sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    sSql = sSql & " Where dFecha='" & sFecha & "' and RH.cUser='" & psCodUser & "' and (C.cRFA='NOR' Or C.cRFA is NULL) and"
    sSql = sSql & " SubString(CS.cCtaCod,4,2)='" & psCodAgen & "' and SubString(CS.cCtaCod,9,1)=" & pnMoneda
    sSql = sSql & " and CS.nDiasAtraso=31 and CS.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ObtenerMontoIngresadoXAnalista = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerMontoPagadoXAnalista(ByVal psCodAgen As String, ByVal pnMoneda As Integer, _
ByVal psCodUser As String, ByVal pdFechaInicial As Date, ByVal pdFechaFinal As Date, _
ByVal pnTipoCambio As Double) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    Dim sFechaInicial As String
    Dim sFechaFinal As String
    
    sFechaInicial = Format(pdFechaInicial, "YYYYMMDD")
    sFechaFinal = Format(pdFechaFinal, "YYYYMMDD")
    
    sSql = "Select Isnull(Sum(Case When SubString(MD.cCtaCod,9,1)='1' Then MD.nMonto Else MD.nMonto*" & pnTipoCambio & " End),0) as nMonto"
    sSql = sSql & " From Mov M"
    sSql = sSql & " Inner Join MovcolDet MD on MD.nMovNro=M.nMovNro"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=MD.cCtaCod and PP.nPrdPersRelac=28"
    sSql = sSql & " Inner Join RRHH  RH on RH.cPersCod=PP.cPersCod"
    sSql = sSql & " Where SubString(MD.cCtaCod,4,2)='" & psCodAgen & "' and SubString(MD.cCtaCod,9,1)='" & pnMoneda & "' and RH.cUser='" & psCodUser & "' and"
    sSql = sSql & " (M.nMovFlag=0 or M.nMovFlag=5) and MD.nPrdConceptoCod=1000 and"
    sSql = sSql & " MD.cOpeCod not in ('100101','100102','100103','100104','100105') and MD.cOpeCod not like '107[123456789]%' and"
    sSql = sSql & " (Left(M.cMovNro,8) Between '" & sFechaInicial & "' and '" & sFechaFinal & "')"
       
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ObtenerMontoPagadoXAnalista = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ObtenerMontosTotalesXAnalista(ByVal pdFecha As Date, ByVal psCodAgen As String, _
ByVal psMoneda As String, ByVal pnTipoCambio As Double, ByVal psCodUser As String) As Recordset
    Dim oConec As DConecta
    Dim sSql As String
    Dim sFechaInicial As String
    
    sFechaInicial = Format(pdFecha, "MM/dd/YYYY")
    
    sSql = "Select Sum(Case When SubString(CS.cCtaCod,9,1)='1' Then CS.nSaldoCap Else CS.nSaldoCap*" & pnTipoCambio & " End ) as nMonto"
    sSql = sSql & " From ColocacSaldo CS"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
    sSql = sSql & " Inner Join RRHH RH on RH.cPersCod=PP.cPersCod"
    sSql = sSql & " Inner Join ColocacCred C on C.cCtaCod=CS.cCtaCod"
    sSql = sSql & " Where SubString(CS.cCtacod,4,2)='" & psCodAgen & "' and CS.dFecha='" & sFechaInicial & "'  and"
    sSql = sSql & " (C.cRFA='NOR' or C.cRFA Is Null) and RH.cUser='" & psCodUser & "'"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set ObtenerMontosTotalesXAnalista = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ObtenerMontosFinalesXAnalista(ByVal pdFecha As Date, ByVal psCodAgen As String, _
ByVal psMoneda As String, ByVal pnTipoCambio As Double, ByVal psCodUser As String) As Recordset

    Dim oConec As DConecta
    Dim sSql As String
    Dim sFechaInicial As String
    
    sFechaInicial = Format(pdFecha, "MM/dd/yyyy")
    
    sSql = "Select Sum(Case When SubString(CS.cCtaCod,9,1)='1' Then CS.nSaldoCap Else CS.nSaldoCap End ) as nMonto"
    sSql = sSql & " From ColocacSaldo CS"
    sSql = sSql & " Inner Join ProductoPersona PP on PP.cCtaCod=CS.cCtaCod and PP.nPrdPersRelac=28"
    sSql = sSql & " Inner Join RrHh RH on RH.cPersCod=PP.cPersCod"
    sSql = sSql & " Inner Join ColocacCred C on C.cCtaCod=CS.cCtaCod"
    sSql = sSql & " Where SubString(CS.cCtacod,4,2)='" & psCodAgen & "' and CS.dFecha='" & sFechaInicial & "' and"
    sSql = sSql & " (C.cRFA='NOR' or C.cRFA Is Null) and RH.cUser='" & psCodUser & "'  and CS.nDiasAtraso>=31 and"
    sSql = sSql & " CS.nPrdEstado in (2020,2021,2022,2030,2031,2032) and SubString(CS.cCtaCod,9,1)='" & psMoneda & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ObtenerMontosFinalesXAnalista = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function ListaCreditosXAnalistaXAgencia(ByVal psCodAgen As String, ByVal psCodAnalista As String, _
ByVal psCodMoneda As Integer) As Recordset
    Dim sSql As String
    Dim oConec As DConecta
    
    sSql = "Select P.cCtaCod,"
    sSql = sSql & " Pers.cPersNombre as cTitular,"
    sSql = sSql & " Pers.cPersTelefono as cTitularTelefono1,"
    sSql = sSql & " Pers.cPersTelefono2 as cTitularTelefono2,"
    sSql = sSql & " Pers.cPersDireccDomicilio as cTitularDomicilio,"
    sSql = sSql & " cTitularDirecNegocio=(Select PFI.cRazSocDirecc"
    sSql = sSql & "                         From ColocFteIngreso CF"
    sSql = sSql & "                                 Inner Join PersFteIngreso PFI on CF.cNumFuente=PFI.cNumFuente"
    sSql = sSql & "                           Where CF.cCtaCod=P.cCtaCod),"
    sSql = sSql & " Pers1.cPersNombre as cNombreAval,"
    sSql = sSql & " Pers1.cPersDireccDomicilio as cAvalDomicilio,"
    sSql = sSql & " Pers.cPersTelefono as cAvalTelefono1,"
    sSql = sSql & " Pers1.cPersTelefono2 as cAvalTelefono2,"
    sSql = sSql & " P.nSaldo as nSK,"
    sSql = sSql & " CR.nSaldoIntComp as nMontoComp,"
    sSql = sSql & " CR.nSaldoIntMor as nMontoMor,"
    sSql = sSql & " CR.nSaldoGasto as nGasto,"
    sSql = sSql & " FechaPago=(Select SubString(cMovNro,7,2)+'/'+SubString(cMovNro,5,2)+'/'+SubString(cMovNro,1,4)"
    sSql = sSql & "             From Mov Where nMovNro=(Select Max(nMovnro) From MovCol Where cCtacod=P.cCtaCod)),"
    sSql = sSql & " MontoPago=(Select Sum(nMonto)"
    sSql = sSql & "            From MovCol Where nMovNro=(Select Max(nMovNro) From MovCol Where cCtaCod=P.cCtaCod)),"
    sSql = sSql & " SubString(P.cCtaCod,4,2) as cAgencia,"
    sSql = sSql & " SubString(P.cCtaCod,9,1) as cMoneda,"
    sSql = sSql & " PP2.cPersCod  as cAnalista"
    sSql = sSql & " From Producto P"
    sSql = sSql & " Inner Join ProductoPersona PP on P.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=20"
    sSql = sSql & " Inner Join Persona Pers on Pers.cPersCod=PP.cPersCod"
    sSql = sSql & " Left Join ProductoPersona PP1 on PP1.cCtacod=P.cCtaCod and PP1.nPrdPersRelac=21"
    sSql = sSql & " Left Join Persona Pers1 on Pers1.cPersCod=PP1.cPersCod"
    sSql = sSql & " Inner Join ColocRecup  CR on CR.cCtaCod=P.cCtaCod"
    sSql = sSql & " Inner Join ProductoPersona PP2 on PP2.cCtaCod=P.cCtaCod and PP2.nPrdPersRelac=28"
    sSql = sSql & " Where P.nPrdEstado in (2201,2205) and SubString(P.cCtaCod,4,2) In (" & psCodAgen & ") and PP2.cPersCod in (" & psCodAnalista & ") and SubString(P.cCtaCod,9,1)= " & psCodMoneda
    sSql = sSql & " Order By PP2.cPersCod,SubString(P.cCtaCod,4,2),SubString(P.cCtaCod,9,1),P.cCtaCod"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ListaCreditosXAnalistaXAgencia = oConec.CargaRecordSet(sSql)
    oConec.CierraConexion
    Set oConec = Nothing
End Function
