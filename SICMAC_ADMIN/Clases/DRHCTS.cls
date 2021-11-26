VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRHCTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Function ActualizaMesProvison(ByVal PersCod As String, _
ByVal RHCod As String, ByVal User As String, ByVal dFecReg As String, _
ByVal Periodo As String, ByVal Estado As String, ByVal Cantidad As Double, _
Optional ActMes As Integer = 1, Optional pnProvision As Currency) As Integer

Dim Sql1 As String
Dim Sql2 As String
Dim Co As DConecta

If ActMes = 1 Then
Sql1 = " Update RHempleado set nRHMesCts =  nRHMesCTS + 1 where cPersCod = '" & Trim(PersCod) & "'"
End If

Set Co = New DConecta

Sql2 = " Insert MovCts"
Sql2 = Sql2 & " (cRHCod,cUser,dFecReg,cPeriodo,cEstado,nCantidad,nProvision)"
Sql2 = Sql2 & " Values"
Sql2 = Sql2 & " ('" & RHCod & "','" & User & "','" & dFecReg & "','" & Periodo & "','" & Estado & "'," & Cantidad & "," & pnProvision & ")"

On Error GoTo CTSErr
    Co.AbreConexion
    Co.BeginTrans
    If ActMes = 1 Then
        Co.Ejecutar (Sql1)
    End If
    Co.Ejecutar (Sql2)
    Co.CommitTrans
    Co.CierraConexion
    ActualizaMesProvison = 1
Set Co = Nothing
Exit Function

CTSErr:
    Co.RollBackTrans
    ActualizaMesProvison = -1
End Function
Function AbonaCTSPasado(ByVal User As String, ByVal dFecReg As String, _
ByVal Periodo As String) As Integer
Dim Sql0 As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Co As DConecta
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Sql0 = "Select Count(*) nro from RHCTSTemp"

Sql1 = "Select RHCTS.cPersCod, nCantidad, cRHCod from RHCTSTemp RHCTS"
Sql1 = Sql1 & " Inner Join RRHH RH on RH.cPersCod = RHCTS.cPersCod"

Sql3 = "Delete RHCTSTemp"

On Error GoTo CTSErr
    Set Co = New DConecta
    Co.AbreConexion
    Set Rs2 = Co.CargaRecordSet(Sql0)
    
    If Rs2!Nro = 0 Then
        abinactspasado = 1
        Co.CierraConexion
        Set Co = Nothing
        Set Rs = Nothing
        Set Rs2 = Nothing
        Exit Function
    Else
        Set Rs = Co.CargaRecordSet(Sql1)
        If Not (Rs.EOF And Rs.BOF) Then
            While Not Rs.EOF
                Sql2 = "Update RHEmpleado set nRHMesCTS = nRHMesCts +" & Rs!nCantidad
                Sql2 = Sql2 & " where cPersCod ='" & Rs!cPersCod & "'"
                
                Sql4 = " Insert MovCts"
                Sql4 = Sql4 & " (cRHCod,cUser,dFecReg,cPeriodo,cEstado,nCantidad)"
                Sql4 = Sql4 & " Values"
                Sql4 = Sql4 & " ('" & Rs!cRHCod & "','" & User & "','" & FecReg & "','" & Periodo & "','" & 3 & "'," & Rs!nCantidad & ")"
                
                Co.Ejecutar (Sql2)
                Co.Ejecutar (Sql4)
                Rs.MoveNext
            Wend
            Co.Ejecutar (Sql3)
        End If
        Co.CierraConexion
        Set Rs = Nothing
        Set Rs2 = Nothing
        Set Co = Nothing
        AbonaCTSPasado = 1
        Exit Function
    End If
CTSErr:
    Co.RollBackTrans
    AbonaCTSPasado = -1
    Set Co = Nothing
    Set Rs = Nothing
    Set Rs2 = Nothing
End Function


Function InsertaCTSTemp(ByVal RHCod As String, ByVal Cant As Double, ByVal Periodo As String, ByVal Fecha As String) As Integer
Dim Sql As String
Dim Sql2 As String
Dim Co As DConecta

Set Co = New DConecta

Sql = "Insert rhctstemp"
Sql = Sql & " (cPersCod,nCantidad,cPeriodo,dFecReg)"
Sql = Sql & " Values"
Sql = Sql & "('" & RHCod & "'," & Cant & ",'" & Periodo & "','" & Fecha & "')"


'Sql2 = " Update RHempleado set nRHMesCts =  nRHMesCTS + " & Cant & " where cPersCod = '" & Trim(RHCod) & "'"
On Error GoTo CTSErr
    Co.AbreConexion
    Co.BeginTrans
    Co.Ejecutar (Sql)
    'Co.Ejecutar (Sql2)
    Co.CommitTrans
    Co.CierraConexion
    InsertaCTSTemp = 1
    Set Co = Nothing
    Exit Function
CTSErr:
    Co.RollBackTrans
    InsertaCTSTemp = -1
    Set Co = Nothing
End Function
Function VerificaProvCTSMes(ByVal psPeriodo As String) As Boolean
Dim Sql As String
Dim Co As DConecta
Dim Rs As ADODB.Recordset
Set Co = New DConecta

Sql = "select count(*) nro from movcts where cPeriodo = '" & psPeriodo & "'"
Sql = Sql & " and cFlag is null and cestado in ('1','2')and cRHCod <> 'XXXXXX'"

Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
If Rs!Nro = 0 Then
    VerificaProvCTSMes = False
Else
    VerificaProvCTSMes = True
End If

Co.CierraConexion
Set Co = Nothing
Set Rs = Nothing
End Function
Function VerificaNuevoPeriodo(ByVal psPeriodo As String) As Boolean
Dim Sql As String
Dim Co As DConecta
Dim Rs As ADODB.Recordset
Set Co = New DConecta

Sql = "select * from MovCts"
Sql = Sql & " where cRHCod ='XXXXXX' and cPeriodo = '" & psPeriodo & "'"
Sql = Sql & " and cEstado = '0' and cFlag is null"
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
If Not (Rs.EOF And Rs.BOF) Then
    VerificaNuevoPeriodo = True
Else
    VerificaNuevoPeriodo = False
End If
Co.CierraConexion
Set Co = Nothing
Set Rs = Nothing
End Function
Sub NuevoPeriodo(ByVal psPeriodo As String, ByVal psFecha As String, _
ByVal psUser)
Dim Sql As String
Dim Co As DConecta
Set Co = New DConecta

'On Error GoTo CTSErr
    Sql = "Insert MovCts"
    Sql = Sql & " (cRHCod,cUser,dFecReg,cPeriodo,cEstado,nCantidad) values"
    Sql = Sql & " ('XXXXXX','" & psUser & "','" & psFecha & "','" & psPeriodo & "','0',0)"
    
    Co.AbreConexion
    Co.BeginTrans
    Co.Ejecutar (Sql)
    Co.Ejecutar ("Update RHEmpleado set nRHMesCTS = 0")
    Co.CommitTrans
    Co.CierraConexion
    Set Co = Nothing
    Exit Sub
'CTSErr:
'    Co.RollBackTrans
'    Set Co = Nothing
End Sub
Function ProvisionadoCTS(ByVal PeriodoCTS As String) As Integer
Dim Co As DConecta
Dim Sql As String
Dim Rs As ADODB.Recordset

Set Co = New DConecta
Sql = "SELECT COUNT(*) NRO FROM MOVCTS WHERE cPeriodo='" & PeriodoCTS & "' and cEstado<>0 and cRHCOD='XXXXXX' and CFLAG IS NULL"
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
Co.CierraConexion
ProvisionadoCTS = 0

If Not (Rs.EOF And Rs.BOF) Then
    ProvisionadoCTS = Rs!Nro
End If

Set Rs = Nothing
Set Co = Nothing
End Function

'Function Carga_CTS_Mes(ByVal psFechaMEs As String, ByVal psFechaGrati As String, ByVal psFecha As String, _
'Optional pbSumado As Boolean = False, Optional pnMesAbonoCts As Integer = 0, _
'Optional psYYYYMMAbonoCts As String = "") As ADODB.Recordset
'Dim Co As DConecta
'Dim sql As String
'Dim rs As ADODB.Recordset
'Set Co = New DConecta
'
'
'sql = " Declare @pdFecha datetime "
'sql = sql & " Declare @lsFecha varchar(8) "
'sql = sql & " set @pdFecha = '" & Format(psFecha, "YYYYMMDD") & "' "
'sql = sql & " Set @lsFecha = Convert(varchar(8),@pdFecha,112) "
'
'If pnMesAbonoCts = 1 Then
''***********
'    If pbSumado Then
'        sql = sql & " SELECT "
'        sql = sql & " Sum((((SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0)) / 12 ) * Isnull(MESCTS,0)) TOTALDEP , cAgenciaAsig, dbo.GetRHTpoContrato (cPersCod, '" & Format(CDate(psFecha), gsFormatoMovFecha) & "') Cont "
'    Else
'        sql = sql & " SELECT dbo.GetRHTpoContrato (cPersCod, '" & Format(CDate(psFecha), gsFormatoMovFecha) & "') Cont,"
'        sql = sql & " CRHCOD,CPERSCOD, CPERSNOMBRE, DINGRESO, SUELDO - ANT_AFP - INCRE_AFP3   REM_ANT_AFP,"
'        sql = sql & " (SUELDO - Isnull(ANT_AFP,0) ) Total, INCRE_AFP3, GRATI, GRATI6,"
'        sql = sql & " (SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0) REMUNERA_IND, MESCTS,"
'        sql = sql & " (((SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0)) / 12 ) * Isnull(MESCTS,0) TOTALDEP, cAgenciaAsig,MontoCTS "
'    End If
'    sql = sql & " FROM (  SELECT AAA.CPERSCOD, CPERSNOMBRE, DINGRESO,"
'
'    sql = sql & " SUELDO =  (select sum(IsNull(nRHConceptoValor,0)) Monto from rhconcepto "
'    sql = sql & "           where CPERSCOD = AAA.CPERSCOD and crhplanillacod = 'E01'"
'    sql = sql & "           and  cRHConceptoCod like '1%'),"
'
'    sql = sql & "         ANT_AFP,INCRE_AFP3, GRATI , GRATI / 6 GRATI6,CRHCOD,MESCTS, cAgenciaAsig,MontoCTS"
'    sql = sql & "     FROM ("
'
'    sql = sql & "         SELECT RH.CPERSCOD,CRHCOD, CPERSNOMBRE, RH.DINGRESO,"
'
'    'Sql = Sql & "         SUM(CASE WHEN  CRHCONCEPTOCOD = '109' THEN NMONTO ELSE 0 END) ANT_AFP,"
'    'Sql = Sql & "         SUM(CASE WHEN  CRHCONCEPTOCOD = '110' THEN NMONTO ELSE 0 END) INCRE_AFP3,"
'    sql = sql & "         SUM(CASE WHEN  crhplanillacod = '109' and crhplanillacod = 'E01' THEN nRHConceptoValor ELSE 0 END) ANT_AFP,"
'    sql = sql & "         SUM(CASE WHEN  crhplanillacod = '110' and crhplanillacod = 'E01' THEN nRHConceptoValor ELSE 0 END) INCRE_AFP3,"
'
'    'cambio Aqui
'    sql = sql & "        sum(case when PLA.crhconceptocod = '130' then Pla.nMonto else 0 end)  / count(*) MontoCTS,"
'
'    sql = sql & "         AVG(RHE.NRHMESCTS) MESCTS,cAgenciaAsig"
'
'    'Sql = Sql & "         FROM RHPLANILLADETCON RHPD"
'    sql = sql & "         FROM rhconcepto RHPD"
'    sql = sql & "         INNER JOIN RRHH RH ON RH.CPERSCOD = RHPD.CPERSCOD " 'and nRHEstado < 700 AND LEFT(RH.cRHCod,1)='E'"
'    sql = sql & "         INNER JOIN PERSONA P ON P.CPERSCOD = RH.CPERSCOD"
'    'cambio Aqui
'    sql = sql & "         inner join rhplanilladetcon PLA on PLA.CPERSCOD = RH.CPERSCOD         "
'    sql = sql & "         INNER JOIN RHEMPLEADO RHE ON RHE.CPERSCOD = RH.CPERSCOD"
'
'    'Sql = Sql & "        WHERE CRRHHPERIODO LIKE '" & psFechaMEs & "%' AND CRHCONCEPTOCOD IN ('110','109','130') and subString(cPlanillaCod,1,1)='E'"
'    'Sql = Sql & "        WHERE CRRHHPERIODO LIKE '" & psFechaMEs & "%' AND CRHCONCEPTOCOD IN ('110','109','130') and subString(cPlanillaCod,1,1)='E'"
'    sql = sql & "         Inner Join     ("
'    sql = sql & "             Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like  Left(@lsFecha,6) +  '%' And cPlanillaCod = 'E01' and cperscod not in (Select cPersCod from rrhh Where datediff(month,dcese,@pdFecha) = 0 and datediff(day,dcese,@pdFecha) <> 0 and crhcod like 'E%' and cperscod  in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%'  And cPlanillaCod = 'E01'))"
'    sql = sql & "             Union"
'    sql = sql & "             Select cPersCod from rrhh Where datediff(month,dingreso,@pdFecha) = 0 and dcese is null and crhcod like 'E%' and cperscod not in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%' And cPlanillaCod = 'E01')"
'    sql = sql & "             Union"
'    sql = sql & "             Select cPersCod from rrhh Where datediff(month,dingreso,@pdFecha) = 0 and crhcod like 'E%' and datediff(month,dcese,@pdFecha) = 0 and datediff(day,dcese,@pdFecha) = 0 and cperscod not in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%'  And cPlanillaCod = 'E01')"
'    sql = sql & "             Union"
'    sql = sql & "             Select cPersCod from rrhh Where datediff(month,dcese,@pdFecha) = 0 and datediff(day,dcese,@pdFecha) = 0 and crhcod like 'E%' and cperscod  in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%' And cPlanillaCod = 'E01')"
'    sql = sql & "         ) E on E.cPersCod = RH.cPersCod"
'
'    sql = sql & "         where datediff(month,dIngreso,'" & Format(psFecha, "YYYY/MM/DD") & "')> =0       "
'    'cambio Aqui
'    sql = sql & "         and crrhhperiodo like '" & psYYYYMMAbonoCts & "%__' and cplanillaCod = 'E05' and PLA.crhconceptocod = '130'  "
'    sql = sql & "         GROUP BY  RH.CPERSCOD,CRHCOD, CPERSNOMBRE, DINGRESO,cAgenciaAsig"
'
'    sql = sql & "     ) AAA"
'    sql = sql & "     LEFT JOIN ("
'    sql = sql & "         select CPERSCOD, NMONTO GRATI  from rhplanilladetcon where CRRHHPERIODO like '" & psFechaGrati & "%' "
'    sql = sql & "         and  cPlanillaCod in ('E02') AND CRHCONCEPTOCOD='130'"
'    sql = sql & "     ) BBB ON AAA.CPERSCOD = BBB.CPERSCOD"
'    sql = sql & " ) XXX"
'
'
'    If pbSumado Then
'        sql = sql & " Group by cAgenciaAsig, dbo.GetRHTpoContrato (cPersCod, '" & Format(CDate(psFecha), gsFormatoMovFecha) & "') Order By Cont"
'    Else
'        sql = sql & " ORDER BY CRHCOD"
'    End If
'
''************
'Else
'    If pbSumado Then
'        sql = sql & " SELECT "
'        sql = sql & " Sum((((SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0)) / 12 ) * Isnull(MESCTS,0)) TOTALDEP , cAgenciaAsig, dbo.GetRHTpoContrato (cPersCod, '" & Format(CDate(psFecha), gsFormatoMovFecha) & "') Cont "
'    Else
'        sql = sql & " SELECT dbo.GetRHTpoContrato (cPersCod, '" & Format(CDate(psFecha), gsFormatoMovFecha) & "') Cont,"
'        sql = sql & " CRHCOD,CPERSCOD, CPERSNOMBRE, DINGRESO, SUELDO - ANT_AFP - INCRE_AFP3   REM_ANT_AFP,"
'        sql = sql & " (SUELDO - Isnull(ANT_AFP,0) ) Total, INCRE_AFP3, GRATI, GRATI6,"
'        sql = sql & " (SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0) REMUNERA_IND, MESCTS,"
'        sql = sql & " (((SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0)) / 12 ) * Isnull(MESCTS,0) TOTALDEP, cAgenciaAsig "
'    End If
'    sql = sql & " FROM (  SELECT AAA.CPERSCOD, CPERSNOMBRE, DINGRESO,"
'
'    sql = sql & " SUELDO =  (select sum(IsNull(nRHConceptoValor,0)) Monto from rhconcepto "
'    sql = sql & "           where CPERSCOD = AAA.CPERSCOD and crhplanillacod = 'E01'"
'    sql = sql & "           and  cRHConceptoCod like '1%'),"
'    'Sql = Sql & "         SUELDO = (SELECT NRHSUELDOMONTO"
'    'Sql = Sql & "         FROM RHSUELDO WHERE CPERSCOD=AAA.CPERSCOD AND"
'    'Sql = Sql & "                         DRHSUELDOFECHA = (SELECT MAX(DRHSUELDOFECHA) FROM RHSUELDO"
'    'Sql = Sql & "                                                            WHERE CPERSCOD=AAA.CPERSCOD) ),"
'
'    sql = sql & "         ANT_AFP,INCRE_AFP3, GRATI , GRATI / 6 GRATI6,CRHCOD,MESCTS, cAgenciaAsig"
'    sql = sql & "     FROM ("
'
'    sql = sql & "         SELECT RH.CPERSCOD,CRHCOD, CPERSNOMBRE, RH.DINGRESO,"
'
'    'Sql = Sql & "         SUM(CASE WHEN  CRHCONCEPTOCOD = '109' THEN NMONTO ELSE 0 END) ANT_AFP,"
'    'Sql = Sql & "         SUM(CASE WHEN  CRHCONCEPTOCOD = '110' THEN NMONTO ELSE 0 END) INCRE_AFP3,"
'    sql = sql & "         SUM(CASE WHEN  crhplanillacod = '109' and crhplanillacod = 'E01' THEN nRHConceptoValor ELSE 0 END) ANT_AFP,"
'    sql = sql & "         SUM(CASE WHEN  crhplanillacod = '110' and crhplanillacod = 'E01' THEN nRHConceptoValor ELSE 0 END) INCRE_AFP3,"
'
'    sql = sql & "         AVG(RHE.NRHMESCTS) MESCTS,cAgenciaAsig"
'
'    'Sql = Sql & "         FROM RHPLANILLADETCON RHPD"
'    sql = sql & "         FROM rhconcepto RHPD"
'    sql = sql & "         INNER JOIN RRHH RH ON RH.CPERSCOD = RHPD.CPERSCOD " 'and nRHEstado < 700 AND LEFT(RH.cRHCod,1)='E'"
'    sql = sql & "         INNER JOIN PERSONA P ON P.CPERSCOD = RH.CPERSCOD"
'    sql = sql & "         INNER JOIN RHEMPLEADO RHE ON RHE.CPERSCOD = RH.CPERSCOD"
'
'    'Sql = Sql & "        WHERE CRRHHPERIODO LIKE '" & psFechaMEs & "%' AND CRHCONCEPTOCOD IN ('110','109','130') and subString(cPlanillaCod,1,1)='E'"
'    'Sql = Sql & "        WHERE CRRHHPERIODO LIKE '" & psFechaMEs & "%' AND CRHCONCEPTOCOD IN ('110','109','130') and subString(cPlanillaCod,1,1)='E'"
'    sql = sql & "         Inner Join     ("
'    sql = sql & "             Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like  Left(@lsFecha,6) +  '%' And cPlanillaCod = 'E01' and cperscod not in (Select cPersCod from rrhh Where datediff(month,dcese,@pdFecha) = 0 and datediff(day,dcese,@pdFecha) <> 0 and crhcod like 'E%' and cperscod  in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%'  And cPlanillaCod = 'E01'))"
'    sql = sql & "             Union"
'    sql = sql & "             Select cPersCod from rrhh Where datediff(month,dingreso,@pdFecha) = 0 and dcese is null and crhcod like 'E%' and cperscod not in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%' And cPlanillaCod = 'E01')"
'    sql = sql & "             Union"
'    sql = sql & "             Select cPersCod from rrhh Where datediff(month,dingreso,@pdFecha) = 0 and crhcod like 'E%' and datediff(month,dcese,@pdFecha) = 0 and datediff(day,dcese,@pdFecha) = 0 and cperscod not in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%'  And cPlanillaCod = 'E01')"
'    sql = sql & "             Union"
'    sql = sql & "             Select cPersCod from rrhh Where datediff(month,dcese,@pdFecha) = 0 and datediff(day,dcese,@pdFecha) = 0 and crhcod like 'E%' and cperscod  in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%' And cPlanillaCod = 'E01')"
'    sql = sql & "         ) E on E.cPersCod = RH.cPersCod"
'
'    sql = sql & "         where datediff(month,dIngreso,'" & Format(psFecha, "YYYY/MM/DD") & "')> =0       "
'    sql = sql & "         GROUP BY  RH.CPERSCOD,CRHCOD, CPERSNOMBRE, DINGRESO,cAgenciaAsig"
'
'    'Sql = Sql & "        Union"
'    'Sql = Sql & "        Select RH.cPersCod, cRHCod, CPERSNOMBRE, dIngreso, 0 ANT_AFP, 0 INCRE_AFP3, IsNull(RHE.NRHMESCTS,0) MESCTS,cAgenciaAsig"
'    'Sql = Sql & "        from RRHH RH"
'    'Sql = Sql & "        INNER JOIN PERSONA P ON P.CPERSCOD = RH.CPERSCOD"
'    'Sql = Sql & "        INNER JOIN RHEMPLEADO RHE ON RHE.CPERSCOD = RH.CPERSCOD"
'    'Sql = Sql & "        where datediff(month,dIngreso,'" & Format(psFecha, "YYYY/MM/DD") & "')=0 and nRHEstado < 700 and cRHCod like 'E%'"
'    'Sql = Sql & "        and not RH.cPersCod in ( Select distinct cPersCod from RHPLANILLADETCON where  CRRHHPERIODO LIKE '" & psFechaMEs & "%' and subString(cPlanillaCod,1,1)='E' )"
'
'    sql = sql & "     ) AAA"
'    sql = sql & "     LEFT JOIN ("
'    sql = sql & "         select CPERSCOD, NMONTO GRATI  from rhplanilladetcon where CRRHHPERIODO like '" & psFechaGrati & "%' "
'    sql = sql & "         and  cPlanillaCod in ('E02') AND CRHCONCEPTOCOD='130'"
'    sql = sql & "     ) BBB ON AAA.CPERSCOD = BBB.CPERSCOD"
'    sql = sql & " ) XXX"
'    '***********
'    'sql = sql & " UNION "
'    'sql = sql & " Select"
'    'If pbSumado = False Then
'    '    sql = sql & " cRHCod, cPersCod,"
'    '    sql = sql & " (Select P.cPersNombre from Persona P where P.cPersCod = ZZZ.cPersCod) cPersNombre,"
'    '    sql = sql & "  dIngreso, ( Sueldo - ANT_AFP - INCRE_AFP3 )  REM_ANT_AFP,"
'    '    sql = sql & " (SUELDO - Isnull(ANT_AFP,0) ) Total, INCRE_AFP3, GRATI , GRATI6,"
'    '    sql = sql & " (SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0) REMUNERA_IND, MESCTS,"
'    '    sql = sql & " (((SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0)) / 12 ) * Isnull(MESCTS,0) TOTALDEP, cAgenciaAsig"
'    'Else
'    '    sql = sql & "  cPersCod,"
'    '    sql = sql & " (Select P.cPersNombre from Persona P where P.cPersCod = ZZZ.cPersCod) cPersNombre,"
'    '    sql = sql & " dIngreso, SUELDO, (Sueldo - ANT_AFP - INCRE_AFP3 )  REM_ANT_AFP,"
'    '    sql = sql & " Isnull(ANT_AFP,0) ANT_AFP,INCRE_AFP3,GRATI , GRATI6,cRHCod,MESCTS, cAgenciaAsig"
'    'End If
'    'sql = sql & " from ("
'    'sql = sql & "     Select AAA.cPersCod, cRHCod,nRHEstado,cAgenciaAsig,dIngreso,"
'    'sql = sql & "     SUELDO = (SELECT NRHSUELDOMONTO FROM RHSUELDO WHERE CPERSCOD=AAA.CPERSCOD AND"
'    'sql = sql & "             DRHSUELDOFECHA = (SELECT MAX(DRHSUELDOFECHA) FROM RHSUELDO WHERE CPERSCOD=AAA.CPERSCOD) ),"
'    'sql = sql & "     Isnull((select IsNull(nRHConceptoValor,0) Valor from RhConcepto where cRHConceptoCod = '109' and cPersCod = AAA.CPERSCOD),0) ANT_AFP,"
'    'sql = sql & "     Isnull((select IsNull(nRHConceptoValor,0) Valor from RhConcepto where cRHConceptoCod = '110' and cPersCod = AAA.CPERSCOD),0) INCRE_AFP3,GRATI, GRATI / 6 GRATI6,MESCTS"
'    'sql = sql & "     from ("
'    'sql = sql & "         select R.cPersCod, cRHCod,nRHEstado,cAgenciaAsig,dIngreso,nRHMESCTS MESCTS"
'    'sql = sql & "         from rrhh R"
'    'sql = sql & "         Inner join RHEmpleado E on R.cPersCod = E.cPersCod"
'    'sql = sql & "         where nRHEstado <700 and datediff(month,dIngreso,'" & Format(psFecha, "YYYY/MM/DD") & "' ) > 0"
'    'sql = sql & "         and substring(cRHCod,1,1)='E' and"
'    'sql = sql & "         not R.cPerscod in"
'    'sql = sql & "         (Select distinct cPersCod from RHPlanillaDetCon where cRRHHPeriodo like '" & psFechaMEs & "%' and substring(cPlanillaCod,1,1)='E')"
'    'sql = sql & "     ) AAA"
'    'sql = sql & "     Left Join"
'    'sql = sql & "     ("
'    'sql = sql & "         select CPERSCOD, NMONTO GRATI   from rhplanilladetcon where CRRHHPERIODO"
'    'sql = sql & "          like '" & psFechaGrati & "%'"
'    'sql = sql & "         and  cPlanillaCod in ('E02') AND CRHCONCEPTOCOD='130'"
'    'sql = sql & "     ) BBB on AAA.cPersCod =  BBB.cPersCod"
'    'sql = sql & " ) ZZZ"
'
'
'    If pbSumado Then
'        sql = sql & " Group by cAgenciaAsig, dbo.GetRHTpoContrato (cPersCod, '" & Format(CDate(psFecha), gsFormatoMovFecha) & "') Order By Cont"
'    Else
'        sql = sql & " ORDER BY CRHCOD"
'    End If
'End If
'Co.AbreConexion
'Set rs = Co.CargaRecordSet(sql)
'Co.CierraConexion
'Set Carga_CTS_Mes = rs
'Set rs = Nothing
'Set Co = Nothing
'
'End Function

Function Carga_CTS_Mes(ByVal psFechaMEs As String, ByVal psFechaGrati As String, ByVal psFecha As String, Optional pbSumado As Boolean = False) As ADODB.Recordset
Dim Co As DConecta
Dim Sql As String
Dim Rs As ADODB.Recordset
Set Co = New DConecta

'If pbSumado Then
'    sql = " SELECT "
'    sql = sql & " Sum((((SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0)) / 12 ) * Isnull(MESCTS,0)) TOTALDEP , cAgenciaAsig, dbo.GetRHTpoContrato (cPersCod, '" & Format(CDate(psFecha), gsFormatoMovFecha) & "') Cont "
'Else
'    sql = " SELECT "
'    sql = sql & " CRHCOD,CPERSCOD, CPERSNOMBRE, DINGRESO, SUELDO - ANT_AFP - INCRE_AFP3   REM_ANT_AFP,"
'    sql = sql & " (SUELDO - Isnull(ANT_AFP,0) ) Total, INCRE_AFP3, GRATI, GRATI6,"
'    sql = sql & " (SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0) REMUNERA_IND, MESCTS,"
'    sql = sql & " (((SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0)) / 12 ) * Isnull(MESCTS,0) TOTALDEP, cAgenciaAsig "
'End If
'sql = sql & " FROM (  SELECT AAA.CPERSCOD, CPERSNOMBRE, DINGRESO,"
'
'sql = sql & " SUELDO =  (select sum(IsNull(nRHConceptoValor,0)) Monto from rhconcepto "
'sql = sql & "           where CPERSCOD = AAA.CPERSCOD and crhplanillacod = 'E01'"
'sql = sql & "           and  cRHConceptoCod like '1%'),"
''Sql = Sql & "         SUELDO = (SELECT NRHSUELDOMONTO"
''Sql = Sql & "         FROM RHSUELDO WHERE CPERSCOD=AAA.CPERSCOD AND"
''Sql = Sql & "                         DRHSUELDOFECHA = (SELECT MAX(DRHSUELDOFECHA) FROM RHSUELDO"
''Sql = Sql & "                                                            WHERE CPERSCOD=AAA.CPERSCOD) ),"
'
'sql = sql & "         ANT_AFP,INCRE_AFP3, GRATI , GRATI / 6 GRATI6,CRHCOD,MESCTS, cAgenciaAsig"
'sql = sql & "     FROM ("
'
'sql = sql & "         SELECT RH.CPERSCOD,CRHCOD, CPERSNOMBRE, RH.DINGRESO,"
'
''Sql = Sql & "         SUM(CASE WHEN  CRHCONCEPTOCOD = '109' THEN NMONTO ELSE 0 END) ANT_AFP,"
''Sql = Sql & "         SUM(CASE WHEN  CRHCONCEPTOCOD = '110' THEN NMONTO ELSE 0 END) INCRE_AFP3,"
'sql = sql & "         SUM(CASE WHEN  crhplanillacod = '109' and crhplanillacod = 'E01' THEN nRHConceptoValor ELSE 0 END) ANT_AFP,"
'sql = sql & "         SUM(CASE WHEN  crhplanillacod = '110' and crhplanillacod = 'E01' THEN nRHConceptoValor ELSE 0 END) INCRE_AFP3,"
'
'sql = sql & "         AVG(RHE.NRHMESCTS) MESCTS,cAgenciaAsig"
'
''Sql = Sql & "         FROM RHPLANILLADETCON RHPD"
'sql = sql & "         FROM rhconcepto RHPD"
'sql = sql & "         INNER JOIN RRHH RH ON RH.CPERSCOD = RHPD.CPERSCOD and nRHEstado < 700 AND LEFT(RH.cRHCod,1)='E'"
'sql = sql & "         INNER JOIN PERSONA P ON P.CPERSCOD = RH.CPERSCOD"
'sql = sql & "         INNER JOIN RHEMPLEADO RHE ON RHE.CPERSCOD = RH.CPERSCOD"
''Sql = Sql & "        WHERE CRRHHPERIODO LIKE '" & psFechaMEs & "%' AND CRHCONCEPTOCOD IN ('110','109','130') and subString(cPlanillaCod,1,1)='E'"
''Sql = Sql & "        WHERE CRRHHPERIODO LIKE '" & psFechaMEs & "%' AND CRHCONCEPTOCOD IN ('110','109','130') and subString(cPlanillaCod,1,1)='E'"
'sql = sql & "         GROUP BY  RH.CPERSCOD,CRHCOD, CPERSNOMBRE, DINGRESO,cAgenciaAsig"
'
'sql = sql & "        Union"
'sql = sql & "        Select RH.cPersCod, cRHCod, CPERSNOMBRE, dIngreso, 0 ANT_AFP, 0 INCRE_AFP3, IsNull(RHE.NRHMESCTS,0) MESCTS,cAgenciaAsig"
'sql = sql & "        from RRHH RH"
'sql = sql & "        INNER JOIN PERSONA P ON P.CPERSCOD = RH.CPERSCOD"
'sql = sql & "        INNER JOIN RHEMPLEADO RHE ON RHE.CPERSCOD = RH.CPERSCOD"
'sql = sql & "        where datediff(month,dIngreso,'" & Format(psFecha, "YYYY/MM/DD") & "')=0 and nRHEstado < 700 and cRHCod like 'E%'"
'sql = sql & "        and not RH.cPersCod in ( Select distinct cPersCod from RHPLANILLADETCON where  CRRHHPERIODO LIKE '" & psFechaMEs & "%' and subString(cPlanillaCod,1,1)='E' )"
'
'sql = sql & "     ) AAA"
'sql = sql & "     LEFT JOIN ("
'sql = sql & "         select CPERSCOD, NMONTO GRATI  from rhplanilladetcon where CRRHHPERIODO like '" & psFechaGrati & "%' "
'sql = sql & "         and  cPlanillaCod in ('E02') AND CRHCONCEPTOCOD='130'"
'sql = sql & "     ) BBB ON AAA.CPERSCOD = BBB.CPERSCOD"
'sql = sql & " ) XXX"
''***********
''sql = sql & " UNION "
''sql = sql & " Select"
''If pbSumado = False Then
''    sql = sql & " cRHCod, cPersCod,"
''    sql = sql & " (Select P.cPersNombre from Persona P where P.cPersCod = ZZZ.cPersCod) cPersNombre,"
''    sql = sql & "  dIngreso, ( Sueldo - ANT_AFP - INCRE_AFP3 )  REM_ANT_AFP,"
''    sql = sql & " (SUELDO - Isnull(ANT_AFP,0) ) Total, INCRE_AFP3, GRATI , GRATI6,"
''    sql = sql & " (SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0) REMUNERA_IND, MESCTS,"
''    sql = sql & " (((SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0)) / 12 ) * Isnull(MESCTS,0) TOTALDEP, cAgenciaAsig"
''Else
''    sql = sql & "  cPersCod,"
''    sql = sql & " (Select P.cPersNombre from Persona P where P.cPersCod = ZZZ.cPersCod) cPersNombre,"
''    sql = sql & " dIngreso, SUELDO, (Sueldo - ANT_AFP - INCRE_AFP3 )  REM_ANT_AFP,"
''    sql = sql & " Isnull(ANT_AFP,0) ANT_AFP,INCRE_AFP3,GRATI , GRATI6,cRHCod,MESCTS, cAgenciaAsig"
''End If
''sql = sql & " from ("
''sql = sql & "     Select AAA.cPersCod, cRHCod,nRHEstado,cAgenciaAsig,dIngreso,"
''sql = sql & "     SUELDO = (SELECT NRHSUELDOMONTO FROM RHSUELDO WHERE CPERSCOD=AAA.CPERSCOD AND"
''sql = sql & "             DRHSUELDOFECHA = (SELECT MAX(DRHSUELDOFECHA) FROM RHSUELDO WHERE CPERSCOD=AAA.CPERSCOD) ),"
''sql = sql & "     Isnull((select IsNull(nRHConceptoValor,0) Valor from RhConcepto where cRHConceptoCod = '109' and cPersCod = AAA.CPERSCOD),0) ANT_AFP,"
''sql = sql & "     Isnull((select IsNull(nRHConceptoValor,0) Valor from RhConcepto where cRHConceptoCod = '110' and cPersCod = AAA.CPERSCOD),0) INCRE_AFP3,GRATI, GRATI / 6 GRATI6,MESCTS"
''sql = sql & "     from ("
''sql = sql & "         select R.cPersCod, cRHCod,nRHEstado,cAgenciaAsig,dIngreso,nRHMESCTS MESCTS"
''sql = sql & "         from rrhh R"
''sql = sql & "         Inner join RHEmpleado E on R.cPersCod = E.cPersCod"
''sql = sql & "         where nRHEstado <700 and datediff(month,dIngreso,'" & Format(psFecha, "YYYY/MM/DD") & "' ) > 0"
''sql = sql & "         and substring(cRHCod,1,1)='E' and"
''sql = sql & "         not R.cPerscod in"
''sql = sql & "         (Select distinct cPersCod from RHPlanillaDetCon where cRRHHPeriodo like '" & psFechaMEs & "%' and substring(cPlanillaCod,1,1)='E')"
''sql = sql & "     ) AAA"
''sql = sql & "     Left Join"
''sql = sql & "     ("
''sql = sql & "         select CPERSCOD, NMONTO GRATI   from rhplanilladetcon where CRRHHPERIODO"
''sql = sql & "          like '" & psFechaGrati & "%'"
''sql = sql & "         and  cPlanillaCod in ('E02') AND CRHCONCEPTOCOD='130'"
''sql = sql & "     ) BBB on AAA.cPersCod =  BBB.cPersCod"
''sql = sql & " ) ZZZ"
'
'
'If pbSumado Then
'    sql = sql & " Group by cAgenciaAsig, dbo.GetRHTpoContrato (cPersCod, '" & Format(CDate(psFecha), gsFormatoMovFecha) & "') Order By Cont"
'Else
'    sql = sql & " ORDER BY CRHCOD"
'End If

'MAVM 20110711 ***
If pbSumado Then
    Sql = " SELECT "
    'sql = sql & " ISNULL(Sum(Case When SubString(Convert(varchar (10), dIngreso, 112), 1, 8) = '" & Format(CDate(psFecha), gsFormatoMovFecha) & "' Then MESCTS * (((SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0)) / 12 ) else(((SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0)) / 12 ) End), 0) TOTALDEP, cAgenciaActual, dbo.GetRHTpoContrato (cPersCod, '" & Format(CDate(psFecha), gsFormatoMovFecha) & "') Cont "
    Sql = Sql & " SUM(Round(ISNULL(Case When SubString(Convert(varchar (10), dIngreso, 112), 1, 8) = '" & Format(CDate(psFecha), gsFormatoMovFecha) & "' Then MESCTS * (((SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0)) / 12 ) else(((SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0)) / 12 )End , 0), 2)) TOTALDEP, cAgenciaActual "
Else
    Sql = " SELECT "
    Sql = Sql & " CRHCOD,CPERSCOD, CPERSNOMBRE, DINGRESO, SUELDO - ANT_AFP - INCRE_AFP3 REM_ANT_AFP,"
    Sql = Sql & " (SUELDO - Isnull(ANT_AFP,0) ) Total, INCRE_AFP3, GRATI, GRATI6,"
    Sql = Sql & " (SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0) REMUNERA_IND, MESCTS,"
    Sql = Sql & " Round(Case When SubString(Convert(varchar (10), dIngreso, 112), 1, 4) + SubString(Convert(varchar (10), dIngreso, 112), 5, 2) = '" & Format(CDate(psFecha), gsFormatoMovFecha) & "' Then MESCTS * (((SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0)) / 12 ) else (((SUELDO - IsNull(ANT_AFP,0) ) + Isnull(GRATI6,0)) / 12 ) End, 2) TOTALDEP, cAgenciaActual "
End If
Sql = Sql & " FROM (  SELECT AAA.CPERSCOD, CPERSNOMBRE, DINGRESO,"

Sql = Sql & " SUELDO = IsNull((Select sum(IsNull(nMonto,0)) Monto From RHPlanillaDetCon Where cPersCod = AAA.cPersCod And (cPlanillaCod = ('E01') Or cPlanillaCod = ('E06') or cPlanillaCod = ('E08')) And cRRHHPeriodo Like '" & Mid(Format(CDate(psFecha), gsFormatoMovFecha), 1, 6) & "%' And cRHConceptoCod like '1%' And cRHConceptoCod <> '112' And cRHConceptoCod <> '130') ,0), "
Sql = Sql & "         ANT_AFP,INCRE_AFP3, ISNULL(GRATI,0) GRATI, ISNULL(GRATI / 6, 0) GRATI6, CRHCOD, MESCTS, cAgenciaActual"
Sql = Sql & "     FROM ("

Sql = Sql & "        Select PE.cPersCod, EM.cRHCod, PE.cPersNombre, EM.dIngreso, SUM(CASE WHEN  crhplanillacod = '109' and crhplanillacod = 'E01' THEN nRHConceptoValor ELSE 0 END) ANT_AFP"
Sql = Sql & "        , SUM(CASE WHEN  crhplanillacod = '110' and crhplanillacod = 'E01' THEN nRHConceptoValor ELSE 0 END) INCRE_AFP3"
Sql = Sql & "        , Round(Isnull(nCantidad, 0),4) MESCTS, cAgenciaActual"
Sql = Sql & "        From RHPlanillaDet E Inner Join Persona PE On E.cPersCod = PE.cPersCod"
Sql = Sql & "        Inner Join RRHH EM On EM.cPersCod = E.cPersCod"
Sql = Sql & "        Inner Join RHConcepto RHPD ON EM.CPERSCOD = RHPD.CPERSCOD"
Sql = Sql & "        And E.cRRHHPeriodo LIKE '" & Mid(Format(CDate(psFecha), gsFormatoMovFecha), 1, 6) & "%'"
'Sql = Sql & "        And E.cPlanillaCod in ('E01')"
Sql = Sql & "        And E.cPlanillaCod in ('E01','E08')"
Sql = Sql & "        Left Join MovCts MC on EM.cRHCod = MC.cRHCod And MC.cPeriodo LIKE '" & Mid(Format(CDate(psFecha), gsFormatoMovFecha), 1, 6) & "'"
Sql = Sql & "        GROUP BY PE.CPERSCOD, EM.CRHCOD, PE.CPERSNOMBRE, EM.DINGRESO, cAgenciaActual, nCantidad"


Sql = Sql & "     ) AAA"
Sql = Sql & "     LEFT JOIN ("
Sql = Sql & "         select CPERSCOD, SUM(NMONTO) GRATI  from rhplanilladetcon RP where CRRHHPERIODO like '" & psFechaGrati & "%' "
Sql = Sql & "         and  cPlanillaCod in ('E02') AND CRHCONCEPTOCOD in ('132','185') And cRRHHPeriodo >= (Select Convert(Varchar(10),dIngreso,112) from RRHH Where cPersCod = RP.CPERSCOD) Group By CPERSCOD"
Sql = Sql & "     ) BBB ON AAA.CPERSCOD = BBB.CPERSCOD"
Sql = Sql & " ) XXX"

If pbSumado Then
    'sql = sql & " Group by cAgenciaActual, dbo.GetRHTpoContrato (cPersCod, '" & Format(CDate(psFecha), gsFormatoMovFecha) & "') Order By Cont, cAgenciaActual"
    Sql = Sql & " Group by cAgenciaActual Order By cAgenciaActual"
Else
    Sql = Sql & " ORDER BY CPERSNOMBRE"
End If

Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
Co.CierraConexion
Set Carga_CTS_Mes = Rs
Set Rs = Nothing
Set Co = Nothing
End Function

'MAVM 20110915***
Public Function CargarProvCTSConsol(ByVal psPeriodo As String) As ADODB.Recordset
Dim Sql As String
Dim Co As DConecta
Dim Rs As ADODB.Recordset
Set Co = New DConecta

Sql = "Select"
Sql = Sql & " EM.cRHCod, PE.cPersNombre"
If (Mid(psPeriodo, 5, 2) = "11" Or Mid(psPeriodo, 5, 2) = "12" Or Mid(psPeriodo, 5, 2) = "01" Or Mid(psPeriodo, 5, 2) = "02" Or Mid(psPeriodo, 5, 2) = "03" Or Mid(psPeriodo, 5, 2) = "04") Then
    Sql = Sql & " , Prov11 = IsNull((Select nProvision From MovCTS Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "11'" & " And cRHCod = EM.cRHCod), 0)"
    Sql = Sql & " , Prov12 = IsNull((Select nProvision From MovCTS Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "12'" & " And cRHCod = EM.cRHCod), 0)"
    Sql = Sql & " , Prov01 = IsNull((Select nProvision From MovCTS Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "01'" & " And cRHCod = EM.cRHCod), 0)"
    Sql = Sql & " , Prov02 = IsNull((Select nProvision From MovCTS Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "02'" & " And cRHCod = EM.cRHCod), 0)"
    Sql = Sql & " , Prov03 = IsNull((Select nProvision From MovCTS Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "03'" & " And cRHCod = EM.cRHCod), 0)"
    Sql = Sql & " , Prov04 = IsNull((Select nProvision From MovCTS Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "04'" & " And cRHCod = EM.cRHCod), 0)"
End If

If (Mid(psPeriodo, 5, 2) = "05" Or Mid(psPeriodo, 5, 2) = "06" Or Mid(psPeriodo, 5, 2) = "07" Or Mid(psPeriodo, 5, 2) = "08" Or Mid(psPeriodo, 5, 2) = "09" Or Mid(psPeriodo, 5, 2) = "10") Then
    Sql = Sql & " , Prov05 = IsNull((Select nProvision From MovCTS Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "05'" & " And cRHCod = EM.cRHCod), 0)"
    Sql = Sql & " , Prov06 = IsNull((Select nProvision From MovCTS Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "06'" & " And cRHCod = EM.cRHCod), 0)"
    Sql = Sql & " , Prov07 = IsNull((Select nProvision From MovCTS Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "07'" & " And cRHCod = EM.cRHCod), 0)"
    Sql = Sql & " , Prov08 = IsNull((Select nProvision From MovCTS Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "08'" & " And cRHCod = EM.cRHCod), 0)"
    Sql = Sql & " , Prov09 = IsNull((Select nProvision From MovCTS Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "09'" & " And cRHCod = EM.cRHCod), 0)"
    Sql = Sql & " , Prov10 = IsNull((Select nProvision From MovCTS Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "10'" & " And cRHCod = EM.cRHCod), 0)"
End If

Sql = Sql & " From RHPlanillaDet E Inner Join Persona PE On E.cPersCod = PE.cPersCod Inner Join RRHH EM On EM.cPersCod = E.cPersCod "
Sql = Sql & " Inner Join RHConcepto RHPD ON EM.cPersCod = RHPD.cPersCod"

If (Mid(psPeriodo, 5, 2) = "11" Or Mid(psPeriodo, 5, 2) = "12" Or Mid(psPeriodo, 5, 2) = "01" Or Mid(psPeriodo, 5, 2) = "02" Or Mid(psPeriodo, 5, 2) = "03" Or Mid(psPeriodo, 5, 2) = "04") Then
    If (Mid(psPeriodo, 5, 2) = "11" Or Mid(psPeriodo, 5, 2) = "12") Then
        Sql = Sql & " And SubString(E.cRRHHPeriodo, 1, 6) Between '" & Mid(psPeriodo, 1, 4) & "11" & "' " & "And '" & psPeriodo & "'"
    Else
        Sql = Sql & " And SubString(E.cRRHHPeriodo, 1, 6) Between '" & Mid(psPeriodo, 1, 4) - 1 & "11" & "' " & "And '" & psPeriodo & "'"
    End If
End If
If (Mid(psPeriodo, 5, 2) <> "11" And Mid(psPeriodo, 5, 2) <> "12") Then
    Sql = Sql & " And SubString(E.cRRHHPeriodo, 1, 6) Between '" & Mid(psPeriodo, 1, 4) - 1 & "11" & "' " & "And '" & psPeriodo & "'"
Else
    Sql = Sql & " And SubString(E.cRRHHPeriodo, 1, 6) Between '" & Mid(psPeriodo, 1, 4) & "11" & "' " & "And '" & psPeriodo & "'"
End If

Sql = Sql & " And E.cPlanillaCod in ('E01') Inner Join RHEmpleado MC on PE.cPersCod = MC.cPersCod Inner Join Agencias A on EM.cAgenciaActual = A.cAgeCod"
Sql = Sql & " Inner Join Areas AR on AR.cAreaCod = EM.cAreaCodActual Where Not PE.cPersCod in ('1090101173286','1090101173294')"
Sql = Sql & " GROUP BY cAgeDescripcion, PE.cPersCod, EM.cRHCod, PE.cPersNombre, EM.dIngreso, cAgenciaActual, nRHEmplVacacionesPend, AR.cAreaDescripcion"
Sql = Sql & " Order By cPersNombre"
    
If Co.AbreConexion Then
    Set CargarProvCTSConsol = Co.CargaRecordSet(Sql)
    Co.CierraConexion
End If
Set Rs = Nothing
Set Co = Nothing
End Function
'MAVM 20110915***
