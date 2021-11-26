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
    Dim loIni As ClasIni
    
    Set loIni = New ClasIni
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

Public Function dObtieneDatosCabeceraRecuperacion(ByVal psCtaCod As String) As Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = " SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion AS cEstado, Coloc.cUltimaActualizacion, T1.cConsDescripcion AS cRelacion, UPPER(T2.cConsDescripcion) " & _
        " AS cProducto, UPPER(T3.cConsDescripcion) AS cMoneda, PP.cPersCod, PERS.cPersNombre, COLOC.dVigencia AS Fecha, PersID.cPersIDnro, " & _
        " CR.cMetLiquid, cTipo=case when CR.nTipCj=1 then 'JUDICIAL' ELSE 'EXTRAJUDICIAL' END, " & _
        " sCondicion=case when P.nPrdEstado in('2202', '2204') then 'CASTIGADO' ELSE 'JUDICIAL' END, " & _
        " sSEstado=case when P.nPrdEstado in('2201', '2202') then 'VIGENTE' ELSE 'CANCELADO' END " & _
        " FROM ProductoPersona PP INNER JOIN Producto P ON P.cCtaCod = PP.cCtaCod " & _
        " INNER JOIN Constante T ON P.nPrdEstado = T.nConsValor INNER JOIN Constante T1 ON PP.nPrdPersRelac = T1.nConsValor INNER JOIN " & _
        " Constante T2 ON SUBSTRING(PP.cCtaCod, 6, 3) = CONVERT(Varchar(3), T2.nConsValor) INNER JOIN " & _
        " Constante T3 ON SUBSTRING(PP.cCtaCod, 9, 1) = CONVERT(Varchar(1), T3.nConsValor) INNER JOIN " & _
        " Persona PERS ON PP.cPersCod = PERS.cPersCod INNER JOIN Colocaciones COLOC ON P.cCtaCod = COLOC.cCtaCod INNER JOIN " & _
        " ColocRecup CR ON COLOC.cCtaCod = CR.cCtaCod LEFT OUTER JOIN PersID ON PERS.cPersCod = PersID.cPersCod " & _
        " WHERE (PP.cCtaCod='" & psCtaCod & "') AND (T1.nConsCod = 3002) AND (T.nConsCod = 3001) " & _
        " AND (T2.nConsCod = 1001) AND (T3.nConsCod = 1011) " & _
        " AND (P.cCtaCod LIKE '_____%') AND (P.nPrdEstado IN (2201, 2202, 2203, 2204)) AND (PP.nPrdPersRelac = 20)  "

Set lrs = coConex.CargaRecordSet(lsSQL)
 
Set dObtieneDatosCabeceraRecuperacion = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosCabeceraRecuperacion>>", Err.Description
    
End Function

Public Function dObtieneDatosTotalesRecuperacion(ByVal psCtaCod As String) As Recordset
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
        "   WHERE (ColocRecup.cCtaCod='" & psCtaCod & "') " & _
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
        " FROM Producto P where p.cctacod='" & psCtaCod & "') K "
        
Set lrs = coConex.CargaRecordSet(lsSQL)
 
Set dObtieneDatosTotalesRecuperacion = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosTotalesRecuperacion>>", Err.Description
    
End Function

Public Function dObtieneGastosRecuperacion(ByVal psCtaCod As String) As Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = " SELECT '' as Item, convert(varchar(20),ColocRecupGastos.dAsigna,103) fecha, " & _
        " ProductoConcepto.cDescripcion, ColocRecupGastos.nMonto AS Importe, " & _
        " ColocRecupGastos.cMotivoGasto AS Origen FROM ColocRecupGastos INNER JOIN " & _
        " ProductoConcepto ON ColocRecupGastos.nPrdConceptoCod = ProductoConcepto.nPrdConceptoCod " & _
        " WHERE (ColocRecupGastos.cCtaCod='" & psCtaCod & "') AND ColocRecupGastos.nColocRecGastoEstado not in ( " & gColRecGastoEstEliminado & " ) " & _
        " order by  ColocRecupGastos.dAsigna "
        
Set lrs = coConex.CargaRecordSet(lsSQL)
 
Set dObtieneGastosRecuperacion = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneGastosRecuperacion>>", Err.Description
    
End Function

Public Function dObtieneTotalesGastosRecuperacion(ByVal psCtaCod As String) As Recordset
'Obtiene Datos para Mostrar en reporte de recuperaciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

Set lrs = New Recordset

On Error GoTo dError

lsSQL = " SELECT  Isnull(SUM(nMonto),0) AS Importe, Isnull(SUM(nMontoPagado),0) AS Pagado " & _
        " From dbo.ColocRecupGastos  " & _
        " Where cCtaCod ='" & psCtaCod & "' AND nColocRecGastoEstado not in ( " & gColRecGastoEstEliminado & " ) "
        
Set lrs = coConex.CargaRecordSet(lsSQL)
 
Set dObtieneTotalesGastosRecuperacion = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosContrato>>", Err.Description
    
End Function


Public Function dObtieneListaAmortizaciones(ByVal psCtaCod As String) As Recordset
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
        " WHERE (MDet.cCtaCod='" & psCtaCod & "') AND MDet.cOpeCod like '13%' ORDER BY Mov.cMovNro "

Set lrs = coConex.CargaRecordSet(lsSQL)
 

Set dObtieneListaAmortizaciones = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneListaAmortizaciones>>", Err.Description
    
End Function

Public Function dObtieneDatosExpediente(ByVal psCtaCod As String) As Recordset
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
        " WHERE (ColocRecupExpediente.cCtaCod='" & psCtaCod & "') and  pp1.NPRDPERSRELAC = 30 "
          
Set lrs = coConex.CargaRecordSet(lsSQL)
 
Set dObtieneDatosExpediente = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosExpediente>>", Err.Description

End Function


Public Function dObtieneDatosGenerales(ByVal psCtaCod As String) As Recordset
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
        " WHERE (ProductoPersona.nPrdPersRelac = 28) AND (Colocaciones.cCtaCod = '" & psCtaCod & "')"
 
Set lrs = coConex.CargaRecordSet(lsSQL)
 
Set dObtieneDatosGenerales = lrs
Set lrs = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Cred en Recuperaciones <<dObtieneDatosGenerales>>", Err.Description

End Function

