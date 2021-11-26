VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRHVacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function GetConceptoTablaImp(ByVal psCod As String) As Double
Dim rsS As ADODB.Recordset
Dim sSql As String
Dim oCon As DConecta
Set oCon = New DConecta

sSql = "select cRHConceptoCod, cRHConceptoDescripcion,cRHConceptoFormula "
sSql = sSql & " from rhconceptotabla where  cRHConceptoCod = '" & psCod & "'"
GetConceptoTablaImp = 0
If oCon.AbreConexion Then
    Set rsS = oCon.Ejecutar(sSql)
    GetConceptoTablaImp = rsS!cRHConceptoFormula
Else
    MsgBox "Error de conexion", vbInformation, "AVISO"
End If
   
Set oCon = Nothing
Set rsS = Nothing
End Function

Public Function GetPersonalVacacionesMes(ByVal psPeriodo As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As DConecta
Set oCon = New DConecta

sSql = " Select RHV.cPersCod, cRHCod, cPeriodo,dFecIni,dFecFin,"
sSql = sSql & " datediff(day,dFecIni,dFecFin) Dias,  cComemta cComenta,"
sSql = sSql & " (Select  cPErsNombre from persona where  cPersCod = RHV.cPersCod) Nombres"
sSql = sSql & " from RHVacGozadas RHV where cPeriodo = '" & psPeriodo & "'"
sSql = sSql & " Order by dFecIni"
If oCon.AbreConexion Then
    Set GetPersonalVacacionesMes = oCon.CargaRecordSet(sSql)
End If
    
Set oCon = Nothing
End Function

Public Function GetTrabajadorVacaciones(ByVal psPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim oCon As DConecta
Set oCon = New DConecta

sSql = " Select cPersCod, cRHCod, cPeriodo,dFecIni,dFecFin,"
sSql = sSql & " datediff(day,dFecIni,dFecFin) Dias,  cComemta cComenta,"
sSql = sSql & " (Select  cPErsNombre from persona where  cPersCod = '" & psPersCod & "') Nombres"
sSql = sSql & " from RHVacGozadas where cPersCod = '" & psPersCod & "'"
sSql = sSql & " Order by dFecIni"
If oCon.AbreConexion Then
    Set GetTrabajadorVacaciones = oCon.CargaRecordSet(sSql)
End If
    
Set oCon = Nothing
End Function
Public Function VerificaDespido(ByVal psPersCod As String) As Boolean
Dim sSql1 As String
Dim Rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta

sSql1 = " Select * from rrhh where substring(convert(char(3),nRHEstado),1,1) in ('7','8')"
sSql1 = sSql1 & " and cPersCod = '" & psPersCod & "'"

VerificaDespido = False

On Error GoTo VacacionesErr
    
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSql1)
        If Not (Rs.EOF And Rs.BOF) Then
            VerificaDespido = True
        End If
    End If
    Set Rs = Nothing
    Set oCon = Nothing
    Exit Function

VacacionesErr:
    MsgBox Err.Description, vbCritical
    Set Rs = Nothing
    Set oCon = Nothing

End Function
Public Sub InsertaDiasVacaciones(ByVal psPersCod As String, ByVal psFechaIni As String, _
ByVal psFechaFin As String, ByVal pnDia As Integer, ByVal psComenta As String, _
ByVal psRHEmp As String, ByVal psCodUser As String, ByVal psFecSis As String)

Dim sSql1 As String
Dim sSql2 As String
Dim sSql3 As String

Dim oCon As DConecta
Set oCon = New DConecta

'Actuliza Vacaciones gozadas
sSql1 = "Update  RHEmpleado Set nRHEmplVacacionesPend = nRHEmplVacacionesPend -  " & pnDia
sSql1 = sSql1 & " where cPersCod = '" & psPersCod & "'"
' 1 Abono de Vaciones
' 2 Goce de Vacaciones
'Inserta Movimiento de Vacaciones
sSql2 = " Insert MovVacaciones"
sSql2 = sSql2 & " (cPersCod,cCodVacaciones,dFecha,cUser,nDias) values "
sSql2 = sSql2 & " ('" & psPersCod & "','2','" & psFecSis & "','" & psCodUser & "'," & -1 * pnDia & ")"

'Inserta Tabla Vacaciones
sSql3 = "Insert RHVacGozadas"
sSql3 = sSql3 & " (cPersCod,cRHCod,cPeriodo,dFecIni,dFecFin,cComemta) values"
sSql3 = sSql3 & " ('" & psPersCod & "','" & psRHEmp & "','" & Format(psFechaIni, "YYYYMM") & "',"
sSql3 = sSql3 & " '" & psFechaIni & "','" & psFechaFin & "','" & Trim(pcComenta) & "')"

On Error GoTo VacacionesErr
If oCon.AbreConexion Then
        oCon.BeginTrans
        oCon.Ejecutar (sSql1)
        oCon.Ejecutar (sSql2)
        oCon.Ejecutar (sSql3)
        oCon.CommitTrans
        oCon.CierraConexion
End If
Set Rs = Nothing
Set oCon = Nothing
Exit Sub
VacacionesErr:
    oCon.RollBackTrans
    MsgBox Err.Description, vbCritical
    Set Rs = Nothing
    Set oCon = Nothing
End Sub

'Public Function AbonaVacacionMes(ByVal psPersCod As String, ByVal psFecha As String, _
'ByVal psUser As String, ByVal psFechaIng As String, ByVal psFecSis As String, ByVal psProvi As String) As Integer
'
'Dim sSql As String
'Dim sSql2 As String
''Dim Rs As ADODB.Recordset
'Dim oCon As DConecta
'Dim nResult As Double
'Dim nDias As Integer
'Dim nFeb As Integer
'Set oCon = New DConecta
'
'Dim rs As New ADODB.Recordset
'
''Verifica si ya se abono las vacaciones
'sSql = " select * from movVacaciones where cPersCod = '" & psPersCod & "' "
'sSql = sSql & " and datediff(month,dFecha,'" & psFecha & "')=0"
'sSql = sSql & " and cCodVacaciones = '1'"
'
'On Error GoTo VacacionesErr
'' 1 Abono de Vaciones
'' 2 Goce de Vacaciones
'
'sSql2 = " Insert MovVacaciones"
'sSql2 = sSql2 & " (cPersCod,cCodVacaciones,dFecha,cUser,nDias) values "
'
'If DateDiff("M", psFechaIng, psFecSis) = 0 Then
'    If CInt(Mid(psFechaIng, 1, 2)) = 1 Then
'        nResult = 2.5
'        sSql2 = sSql2 & " ('" & psPersCod & "','1','" & psFecha & "','" & psUser & "'," & nResult & ")"
'    Else
'        'nDias = CInt(Left(DateAdd("D", -1, "01" & Mid(DateAdd("M", 1, psFecSis), 3, 12)), 2)) - CInt(Mid(psFechaIng, 1, 2))
'
'        If Month(psFechaIng) = 2 Then
'            nDias = 30 - CInt(Mid(psFechaIng, 1, 2)) + 3
'        Else
'            nDias = 30 - CInt(Mid(psFechaIng, 1, 2)) + 1
'        End If
'
'        If nDias = 0 Then nDias = 1
'        nResult = Round((nDias / 30) * 2.5, 2)
'        sSql2 = sSql2 & " ('" & psPersCod & "','1','" & psFecha & "','" & psUser & "'," & nResult & ")"
'    End If
'Else
'    nResult = 2.5
'    sSql2 = sSql2 & " ('" & psPersCod & "','1','" & psFecha & "','" & psUser & "'," & nResult & ")"
'End If
'
'
'If oCon.AbreConexion Then
'    Set rs = oCon.CargaRecordSet(sSql)
'    AbonaVacacionMes = 0
'    oCon.BeginTrans
'    If rs.EOF And rs.BOF Then
'        AbonaVacacionMes = 1
'        sSql = "Update  RHEmpleado Set nRHEmplVacacionesPend = nRHEmplVacacionesPend + " & nResult
'        sSql = sSql & " where cPersCod = '" & psPersCod & "'"
'        oCon.Ejecutar (sSql)
'        oCon.Ejecutar (sSql2)
'
'        sSql = "select count(*) Nro from RHControlVac where cPersCod = '" & psPersCod & "' and cPeriodo = '" & Left(psProvi, 4) & "'"
'        Set rs = oCon.Ejecutar(sSql)
'        If rs!Nro = 1 Then
'            sql = "update RHControlVac set nAcumulado = nAcumulado + " & nResult & " where cPersCod = '" & psPersCod & "' and cPeriodo = '" & Left(psProvi, 4) & "'"
'        Else
'            sql = " Insert RHControlVac"
'            sql = sql & " (cPerido,nAcumulado,cPersCod)values"
'            sql = sql & " ('" & Left(psProvi, 4) & "'," & nResult & ",'" & psPersCod & "')"
'        End If
'        oCon.Ejecutar (sSql)
'
'
'        oCon.CommitTrans
'    End If
'    oCon.CierraConexion
'End If
'Set rs = Nothing
'Set oCon = Nothing
'Exit Function
'VacacionesErr:
'    oCon.RollBackTrans
'    AbonaVacacionMes = -1
'    MsgBox Err.Description, vbCritical
'    Set rs = Nothing
'    Set oCon = Nothing
'
'End Function

'Public Function GetProvisionVacaciones(ByVal pnTipo As Integer, ByVal psFecSis As String, _
'Optional pdFechaProv As Date = "2005/01/01") As ADODB.Recordset
'Dim oCon As DConecta
'Dim rsC As ADODB.Recordset
'Dim sSql As String
'Dim psPeriodoAnt As String
'Set rsC = New ADODB.Recordset
'Set oCon = New DConecta
'psPeriodoAnt = Format(DateAdd("m", -1, Format(psFecSis, "YYYY/MM/DD")), "YYYYMM")
'
'sSql = " Declare @pdFecha datetime"
'sSql = sSql & " Declare @lsFecha varchar(8)"
'sSql = sSql & " set @pdFecha = '" & Format(pdFechaProv, "YYYY/MM/DD") & "'"
'sSql = sSql & " Set @lsFecha = Convert(varchar(8),@pdFecha,112) "
'
'sSql = sSql & " Select"
'sSql = sSql & " AAA.cPersCod,cPersNombre,cRHCod,nRHEmplVacacionesPend,dIngreso,cAgenciaAsig,"
'sSql = sSql & " BBB.cPersCod,dRHSueldoFecha,nRHSueldoMonto,"
'
''sSql = sSql & " nBono =(select IsNull(sum(nMonto),0) x from RHPlanillaDetcon where cRHConceptoCod = '105' and cPersCod =  AAA.cPersCod"
''sSql = sSql & " and cRRHHPeriodo like '" & psPeriodoAnt & "%'),"
'sSql = sSql & " nBono = (select sum(IsNull(nRHConceptoValor,0)) Monto from rhconcepto where CPERSCOD = AAA.cPersCod and crhplanillacod = 'E01'and  cRHConceptoCod like '1%') ,"
'sSql = sSql & " CCC.cPersCod , cRHContratoNro, nRHContratoTpo,cagedescripcion,careadescripcion,(select  max(dfecini) from RHVacGozadas where cperscod = aaa.cPersCod)FecIniV"
'sSql = sSql & " ,(select  max(dfecfin) from RHVacGozadas where cperscod = aaa.cPersCod)FecFinV"
'sSql = sSql & " from ("
'
''sSql = sSql & "     select RE.cPersCod,PE.cPersNombre, RH.cRHCod, RE.nRHEmplVacacionesPend, RH.dIngreso,ag.cagedescripcion,a.careadescripcion"
''sSql = sSql & "     from rhempleado RE"
''sSql = sSql & "     Inner Join RRHH RH on RH.cPersCod = RE.cPersCod"
''sSql = sSql & "     inner join agencias ag on cagecod=rh.cagenciaactual"
''sSql = sSql & "     inner join areas a on a.careacod=rh.careacodactual"
''sSql = sSql & "     Inner Join Persona PE on PE.cPersCod = RE.cPersCod"
'
'sSql = sSql & "     Select"
'sSql = sSql & "         RE.cPersCod,PE.cPersNombre, RH.cRHCod, RE.nRHEmplVacacionesPend,RH.cAgenciaAsig,"
'sSql = sSql & "         RH.dIngreso , ag.cAgeDescripcion, a.careadescripcion"
'sSql = sSql & "     From"
'sSql = sSql & "     ("
'sSql = sSql & "         Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like  Left(@lsFecha,6) +  '%' And cPlanillaCod = 'E01' and cperscod not in (Select cPersCod from rrhh Where datediff(month,dcese,@pdFecha) = 0 and datediff(day,dcese,@pdFecha) <> 0 and crhcod like 'E%' and cperscod  in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%'  And cPlanillaCod = 'E01'))"
'sSql = sSql & "         Union"
'sSql = sSql & "         Select cPersCod from rrhh Where datediff(month,dingreso,@pdFecha) = 0 and dcese is null and crhcod like 'E%' and cperscod not in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%' And cPlanillaCod = 'E01')"
'sSql = sSql & "         Union"
'sSql = sSql & "         Select cPersCod from rrhh Where datediff(month,dingreso,@pdFecha) = 0 and crhcod like 'E%' and datediff(month,dcese,@pdFecha) = 0 and datediff(day,dcese,@pdFecha) = 0 and cperscod not in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%'  And cPlanillaCod = 'E01')"
'sSql = sSql & "         Union"
'sSql = sSql & "         Select cPersCod from rrhh Where datediff(month,dcese,@pdFecha) = 0 and datediff(day,dcese,@pdFecha) = 0 and crhcod like 'E%' and cperscod  in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%' And cPlanillaCod = 'E01')"
'sSql = sSql & "     ) X"
'sSql = sSql & "     Inner Join RHEmpleado RE on RE.cPersCod = X.cPersCod"
'sSql = sSql & "     Inner Join RRHH RH on RH.cPersCod = RE.cPersCod"
'sSql = sSql & "     inner join agencias ag on cagecod=rh.cagenciaactual"
'sSql = sSql & "     inner join areas a on a.careacod=rh.careacodactual"
'sSql = sSql & "     Inner Join Persona PE on PE.cPersCod = RE.cPersCod"
'
'
''sSql = sSql & "     where not substring(convert(char(3),nRHEstado),1,1) in ('7','8') and substring(cRHCod,1,1)='E'"
'sSql = sSql & " ) AAA"
'sSql = sSql & " Inner Join ("
'sSql = sSql & "     Select cPersCod,dRHSueldoFecha,nRHSueldoMonto from RHSueldo RS where    dRHSueldoFecha = ("
'sSql = sSql & "     Select max(dRHSueldoFecha) from RHSueldo where  cPersCod = RS.cPersCod)"
'sSql = sSql & " ) BBB on BBB.cPersCod = AAA.cPersCod"
'sSql = sSql & " Inner Join ("
'sSql = sSql & "     Select cPersCod,cRHContratoNro,nRHContratoTpo,dFechaFin from RHContrato RE where    cRHContratoNro = ("
'sSql = sSql & "     Select max(cRHContratoNro) from RHContrato where  cPersCod = RE.cPersCod)"
'sSql = sSql & " ) CCC on CCC.cPersCod = AAA.cPersCod"
'Select Case pnTipo
'    Case 0
'        'TODOS LOS EMPLEADOS
'        'sSql = sSql = "Where CCC.nRHContratoTpo = '0'"
'
'    Case 1
'        'ESTABLEBES
'        sSql = sSql & " Where CCC.nRHContratoTpo = '0' "
'
'    Case 2
'        'CONTRATADOS
'        sSql = sSql & " Where CCC.nRHContratoTpo = '1' "
'
'End Select
'sSql = sSql & " and datediff(month,AAA.dIngreso,'" & Format(psFecSis, "YYYY/MM/DD") & "')<>0"
'sSql = sSql & " Order by AAA.cRHCod "
'
'    If oCon.AbreConexion Then
'        Set rsC = oCon.CargaRecordSet(sSql)
'        Set GetProvisionVacaciones = rsC
'        oCon.CierraConexion
'    End If
'    Set rsC = Nothing
'    Set oCon = Nothing
'End Function

Public Function GetProvisionVacaciones(ByVal pnTipo As Integer, ByVal psFecha As String) As ADODB.Recordset
Dim oCon As DConecta
Dim rsC As ADODB.Recordset
Dim sSql As String
Dim psPeriodoAnt As String
Set rsC = New ADODB.Recordset
Set oCon = New DConecta
'psPeriodoAnt = Format(DateAdd("m", -1, Format(psFecSis, "YYYY/MM/DD")), "YYYYMM") 'Comentado Por MAVM 20110713
psPeriodoAnt = Format(DateAdd("m", -1, Format(psFecha, "YYYY/MM/DD")), "YYYYMM") 'MAVM 20110713

'Comentado Por MAVM 20110713 ***
'sSql = " Select"
'sSql = sSql & " AAA.cPersCod,cPersNombre,cRHCod,nRHEmplVacacionesPend,dIngreso,"
'sSql = sSql & " BBB.cPersCod,dRHSueldoFecha,nRHSueldoMonto,"
'
''sSql = sSql & " nBono =(select IsNull(sum(nMonto),0) x from RHPlanillaDetcon where cRHConceptoCod = '105' and cPersCod =  AAA.cPersCod"
''sSql = sSql & " and cRRHHPeriodo like '" & psPeriodoAnt & "%'),"
'sSql = sSql & " nBono = (select sum(IsNull(nRHConceptoValor,0)) Monto from rhconcepto where CPERSCOD = AAA.cPersCod and crhplanillacod = 'E01'and  cRHConceptoCod like '1%') ,"
'sSql = sSql & " CCC.cPersCod , cRHContratoNro, nRHContratoTpo,cagedescripcion,careadescripcion,(select  max(dfecini) from RHVacGozadas where cperscod = aaa.cPersCod)FecIniV"
'sSql = sSql & " ,(select  max(dfecfin) from RHVacGozadas where cperscod = aaa.cPersCod)FecFinV"
'sSql = sSql & " from ("
'sSql = sSql & "     select RE.cPersCod,PE.cPersNombre, RH.cRHCod, RE.nRHEmplVacacionesPend, RH.dIngreso,ag.cagedescripcion,a.careadescripcion"
'sSql = sSql & "     from rhempleado RE"
'sSql = sSql & "     Inner Join RRHH RH on RH.cPersCod = RE.cPersCod"
'sSql = sSql & "     inner join agencias ag on cagecod=rh.cagenciaactual"
'sSql = sSql & "     inner join areas a on a.careacod=rh.careacodactual"
'sSql = sSql & "     Inner Join Persona PE on PE.cPersCod = RE.cPersCod"
'sSql = sSql & "     where not substring(convert(char(3),nRHEstado),1,1) in ('7','8') and substring(cRHCod,1,1)='E'"
'sSql = sSql & " ) AAA"
'sSql = sSql & " Inner Join ("
'sSql = sSql & "     Select cPersCod,dRHSueldoFecha,nRHSueldoMonto from RHSueldo RS where    dRHSueldoFecha = ("
'sSql = sSql & "     Select max(dRHSueldoFecha) from RHSueldo where  cPersCod = RS.cPersCod)"
'sSql = sSql & " ) BBB on BBB.cPersCod = AAA.cPersCod"
'sSql = sSql & " Inner Join ("
'sSql = sSql & "     Select cPersCod,cRHContratoNro,nRHContratoTpo,dFechaFin from RHContrato RE where    cRHContratoNro = ("
'sSql = sSql & "     Select max(cRHContratoNro) from RHContrato where  cPersCod = RE.cPersCod)"
'sSql = sSql & " ) CCC on CCC.cPersCod = AAA.cPersCod"
'Select Case pnTipo
'    Case 0
'        'TODOS LOS EMPLEADOS
'        'sSql = sSql = "Where CCC.nRHContratoTpo = '0'"
'
'    Case 1
'        'ESTABLEBES
'        sSql = sSql & " Where CCC.nRHContratoTpo = '0' "
'
'    Case 2
'        'CONTRATADOS
'        sSql = sSql & " Where CCC.nRHContratoTpo = '1' "
'
'End Select
'sSql = sSql & " and datediff(month,AAA.dIngreso,'" & Format(psFecSis, "YYYY/MM/DD") & "')<>0"
'sSql = sSql & " Order by AAA.cRHCod "
'***

'MAVM 20110713 ***
sSql = " Select"
sSql = sSql & " AAA.cPersCod, cAgeDescripcion, cPersNombre, cRHCod, nRHEmplVacacionesPend, dIngreso, cAgenciaActual,"
'sSql = sSql & " nBono = (Select sum(IsNull(nMonto,0)) Monto From RHPlanillaDetCon Where cPersCod = AAA.cPersCod And cPlanillaCod in ('E01','E06') And cRRHHPeriodo Like '" & Mid(Format(CDate(psFecha), gsFormatoMovFecha), 1, 6) & "%' And cRHConceptoCod like '1%' And cRHConceptoCod <> '112' And cRHConceptoCod <> '130')"
sSql = sSql & " nBono = nRHSueldoMonto,"
sSql = sSql & " (Round(Cast((nRHEmplVacacionesPend) As Decimal (11, 8)) * Cast(nRHSueldoMonto/ 30.00 As Decimal (11, 8)), 2))nRHVacacionesMonto, nRHrovisionMonto"
sSql = sSql & " from ("
sSql = sSql & "     Select Distinct(PE.cPersCod), cAgeDescripcion, PE.cPersNombre, EM.cRHCod, Isnull(nRHEmplVacacionesPend, 0) nRHEmplVacacionesPend, EM.dIngreso, cAgenciaActual,"
sSql = sSql & "     IsNull((Select sum(IsNull(nMonto,0)) Monto from RHPlanillaDetCon Where cPersCod = PE.cPersCod And (cPlanillaCod = ('E01') Or cPlanillaCod = ('E06') or cPlanillaCod = ('E08')) And cRRHHPeriodo Like '" & Mid(Format(CDate(psFecha), gsFormatoMovFecha), 1, 6) & "%' And cRHConceptoCod like '1%' And cRHConceptoCod <> '112' And cRHConceptoCod <> '130') ,0) nRHSueldoMonto,"
sSql = sSql & "     IsNull((select sum(IsNull(nProvision,0)) Monto from MovVacaciones Where cPersCod = PE.cPersCod and cPeriodo = SubString(E.cRRHHPeriodo, 1, 6)) ,0) nRHrovisionMonto"
sSql = sSql & "     From RHPlanillaDet E Inner Join Persona PE On E.cPersCod = PE.cPersCod"
sSql = sSql & "     Inner Join RRHH EM On EM.cPersCod = E.cPersCod"
sSql = sSql & "     Inner Join RHConcepto RHPD ON EM.cPersCod = RHPD.cPersCod"
sSql = sSql & "     And E.cRRHHPeriodo LIKE '" & Mid(Format(CDate(psFecha), gsFormatoMovFecha), 1, 6) & "%'"
sSql = sSql & "     And E.cPlanillaCod in ('E01','E08')"
sSql = sSql & "     Inner Join RHEmpleado MC on PE.cPersCod = MC.cPersCod Inner Join Agencias A on EM.cAgenciaActual = A.cAgeCod"
sSql = sSql & "     GROUP BY cAgeDescripcion, PE.cPersCod, EM.cRHCod, PE.cPersNombre, EM.dIngreso, cAgenciaActual, nRHEmplVacacionesPend, E.cRRHHPeriodo"
sSql = sSql & " ) AAA"

'MAVM Comentado Por MAVM 20110713 ***
'Select Case pnTipo
'    Case 1
'        'ESTABLEBES
'        sSql = sSql & " Where CCC.nRHContratoTpo = '0' "
'
'    Case 2
'        'CONTRATADOS
'        sSql = sSql & " Where CCC.nRHContratoTpo = '1' "
'
'End Select
'***
sSql = sSql & " Order by AAA.cPersNombre"
'***

    If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sSql)
        Set GetProvisionVacaciones = rsC
        oCon.CierraConexion
    End If
    Set rsC = Nothing
    Set oCon = Nothing
End Function

Public Function GetVacacionesEjecutadas(ByVal psFecha As String) As ADODB.Recordset
Dim Sql As String
Dim Co As DConecta
Set Co = New DConecta

Sql = " Select (Select cRHCod from RRHH RH where RH.cPersCod = PER.cPersCod) Codigo,"
Sql = Sql & " cPersNombre , Total, Quinta, Dias_Vac"
Sql = Sql & " from Persona PER"
Sql = Sql & " Inner Join ("
Sql = Sql & " Select cPlanillaCod, cPersCod,"
Sql = Sql & "     sum(case when cRHConceptoCod = '130' then nMonto else 0 end) Total,"
Sql = Sql & "     sum(case when cRHConceptoCod = '201' then nMonto else 0 end) Quinta,"
Sql = Sql & "     sum(case when cRHConceptoCod = '413' then nMonto else 0 end) Dias_Vac"
Sql = Sql & "     from RHPLANILLADETCON RHP"
Sql = Sql & "     WHERE cRRHHPeriodo LIKE '" & psFecha & "%' AND cPlanillaCod in ('E06')"
Sql = Sql & "     and  cRHConceptoCod in ('130','201','413')"
Sql = Sql & "     Group by  cPlanillaCod, cPersCod"
Sql = Sql & " ) PLA on PLA.cPersCod = PER.cPersCod"
Sql = Sql & " Order by  Codigo"

Co.AbreConexion
Set GetVacacionesEjecutadas = Co.CargaRecordSet(Sql)
Co.CierraConexion
Set Co = Nothing
End Function

Public Function GetProvisionMes(ByVal pdFecha As String, ByVal pnMes As Integer) As ADODB.Recordset
Dim Sql As String
Dim Co As DConecta
Dim Rs As ADODB.Recordset
Set Co = New DConecta

Sql = "Select"
Sql = Sql & " AAA.cPersCod,cPersNombre,cRHCod,nRHEmplVacacionesPend,dIngreso,"
Sql = Sql & " nRHSueldoMonto = (select sum(IsNull(nRHConceptoValor,0)) Monto from rhconcepto where CPERSCOD = AAA.cPersCod and crhplanillacod = 'E01' and  cRHConceptoCod like '1%'),"
'Sql = Sql & " BBB.cPersCod,dRHSueldoFecha,nRHSueldoMonto,"
Sql = Sql & " CCC.cPersCod , cRHContratoNro, nRHContratoTpo, dFechaFin"
Sql = Sql & " from ("
Sql = Sql & "     select RE.cPersCod,PE.cPersNombre, RH.cRHCod, RE.nRHEmplVacacionesPend, RH.dIngreso"
Sql = Sql & "     from rhempleado RE"
Sql = Sql & "     Inner Join RRHH RH on RH.cPersCod = RE.cPersCod"
Sql = Sql & "     Inner Join Persona PE on PE.cPersCod = RE.cPersCod"
Sql = Sql & "     where not substring(convert(char(3),nRHEstado),1,1) in ('7','8') and substring(cRHCod,1,1)='E'"
Sql = Sql & " ) AAA"

'Sql = Sql & " Inner Join ("
'Sql = Sql & "     Select cPersCod,dRHSueldoFecha,nRHSueldoMonto from RHSueldo RS where    dRHSueldoFecha = ("
'Sql = Sql & "     Select max(dRHSueldoFecha) from RHSueldo where  cPersCod = RS.cPersCod)"
'Sql = Sql & " ) BBB on BBB.cPersCod = AAA.cPersCod"

Sql = Sql & " Inner Join ("
Sql = Sql & "     Select cPersCod,cRHContratoNro,nRHContratoTpo,dFechaFin from RHContrato RE where    cRHContratoNro = ("
Sql = Sql & "     Select max(cRHContratoNro) from RHContrato where  cPersCod = RE.cPersCod)"
Sql = Sql & " ) CCC on CCC.cPersCod = AAA.cPersCod"
Sql = Sql & " Where "
Sql = Sql & " month(dIngreso) = " & pnMes & " and datediff(month,dIngreso,'" & Format(pdFecha, "YYYY/MM/DD") & "')<>0"
Sql = Sql & " and nRHEmplVacacionesPend >=30"
Sql = Sql & " Order by AAA.cRHCod"
    
    If Co.AbreConexion Then
        Set GetProvisionMes = Co.CargaRecordSet(Sql)
        Co.CierraConexion
    End If
    Set Rs = Nothing
    Set Co = Nothing
End Function

Public Function AbonaVacacionMes(ByVal psPersCod As String, ByVal psFechaProvVac As String, _
ByVal psUser As String, ByVal psFechaIng As String, ByVal psFecSis As String) As Integer

Dim sSql As String
Dim sSql2 As String
Dim Rs As ADODB.Recordset
Dim oCon As DConecta
Dim nResult As Double
Dim nDias As Integer
Dim nFeb As Integer
Set oCon = New DConecta


'Verifica si ya se abono las vacaciones
'sSql = " select * from movVacaciones where cPersCod = '" & psPersCod & "' "
'sSql = sSql & " and datediff(month,dFecha,'" & psFecha & "')=0"
'sSql = sSql & " and cCodVacaciones = '1'"

On Error GoTo VacacionesErr
' 1 Abono de Vaciones
' 2 Goce de Vacaciones

sSql2 = "Insert MovVacaciones"
'Por MAVM 20110713
'sSql2 = sSql2 & " (cPersCod,cCodVacaciones,dFecha,cUser,nDias) values "
sSql2 = sSql2 & " (cPersCod,cCodVacaciones,dFecha,cUser,nDias, cFlag, cPeriodo) values "
'***

If DateDiff("M", psFechaIng, psFechaProvVac) = 0 Then
    'Comentado Por MAVM 20110713 ***
    'If CInt(Mid(psFechaIng, 1, 2)) = 1 Then
    '    nResult = 2.5
    '    sSql2 = sSql2 & " ('" & psPersCod & "','1','" & psFecha & "','" & psUser & "'," & nResult & ")"
    'Else
    '***
        'nDias = CInt(Left(DateAdd("D", -1, "01" & Mid(DateAdd("M", 1, psFecSis), 3, 12)), 2)) - CInt(Mid(psFechaIng, 1, 2))
        
        'Comentado Por MAVM 20110713 ***
        'If Month(psFechaIng) = 2 Then
            'nDias = 30 - CInt(Mid(psFechaIng, 1, 2)) + 3
        'Else
            'nDias = 30 - CInt(Mid(psFechaIng, 1, 2)) + 1
        'End If
        'nDias = 30 - Mid(psFechaIng, 1, 2) 'DateDiff("d", psFechaIng, psFechaProvVac)
        '***
        
        'Comentado Por MAVM 20110713 ***
        'If nDias = 0 Then nDias = 1
        'nResult = Round((nDias / 30) * 2.5, 2)
        '***
        
        nResult = 2.5 'MAVM 20110713
        
        sSql2 = sSql2 & " ('" & psPersCod & "','1','" & psFecSis & "','" & psUser & "'," & nResult & ", 1,'" & Format(psFechaProvVac, "yyyymm") & "'" & ")"
    'End If 'Comentado Por MAVM 20110713
Else
    nResult = 2.5
    sSql2 = sSql2 & " ('" & psPersCod & "','1','" & psFecSis & "','" & psUser & "'," & nResult & ", 1,'" & Format(psFechaProvVac, "yyyymm") & "'" & ")"
End If


If oCon.AbreConexion Then
    'Comentado Por MAVM 20110713 ***
    'Set Rs = oCon.CargaRecordSet(sSql)
    'AbonaVacacionMes = 0
    'If Rs.EOF And Rs.BOF Then
    '***
        AbonaVacacionMes = 1
        sSql = "Update RHEmpleado Set nRHEmplVacacionesPend = nRHEmplVacacionesPend + " & nResult
        sSql = sSql & " where cPersCod = '" & psPersCod & "'"
        oCon.Ejecutar (sSql)
        oCon.Ejecutar (sSql2)
        oCon.BeginTrans
        oCon.CommitTrans
    'End If 'Comentado Por MAVM 20110713
    oCon.CierraConexion
End If
Set Rs = Nothing
Set oCon = Nothing
Exit Function
VacacionesErr:
    oCon.RollBackTrans
    AbonaVacacionMes = -1
    MsgBox Err.Description, vbCritical
    Set Rs = Nothing
    Set oCon = Nothing
    
End Function

Function AbonoVacaciones(ByVal sPeriodoVac As String) As Integer
Dim Co As DConecta
Dim Sql As String
Dim Rs As ADODB.Recordset

Set Co = New DConecta
Sql = "Select Count(nMovVacaciones) Nro From MovVacaciones Where cPeriodo = '" & Format(sPeriodoVac, "YYYYMM") & "' and cCodVacaciones = '1'"
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
Co.CierraConexion
AbonoVacaciones = 0

If Not (Rs.EOF And Rs.BOF) Then
    AbonoVacaciones = Rs!Nro
End If

Set Rs = Nothing
Set Co = Nothing
End Function

'MAVM 20110715 ***
Public Sub ActualizaProvision(ByVal pnProvision As Currency, ByVal psPersCod As String, ByVal psPeriodo As String)
Dim sSql As String

Dim oCon As DConecta
Set oCon = New DConecta

On Error GoTo ActualizaProvisionErr

sSql = "Update MovVacaciones Set nProvision = " & pnProvision
sSql = sSql & " Where cPersCod = '" & psPersCod & "'" & "And cPeriodo = '" & psPeriodo & "'"

If oCon.AbreConexion Then
    oCon.BeginTrans
    oCon.Ejecutar (sSql)
    oCon.CommitTrans
    oCon.CierraConexion
End If
Set Rs = Nothing
Set oCon = Nothing
Exit Sub
ActualizaProvisionErr:
    oCon.RollBackTrans
    MsgBox Err.Description, vbCritical
    Set Rs = Nothing
    Set oCon = Nothing
End Sub

Public Function CargarProvVacacionesConsol(ByVal psPeriodo As String) As ADODB.Recordset
Dim Sql As String
Dim Co As DConecta
Dim Rs As ADODB.Recordset
Set Co = New DConecta

Sql = "Select"
'Sql = Sql & " PE.cPersCod, cAgenciaActual, EM.cRHCod, PE.cPersNombre, cAgeDescripcion, EM.dIngreso, nRHEmplVacacionesPend, AR.cAreaDescripcion, IsNull((select sum(IsNull(nRHConceptoValor,0)) Monto from rhconcepto where CPERSCOD = PE.cPersCod and crhplanillacod = 'E01' And  cRHConceptoCod like '1%') ,0) nRHSueldoMonto"
Sql = Sql & " PE.cPersCod, cAgenciaActual, EM.cRHCod, PE.cPersNombre, cAgeDescripcion, EM.dIngreso, nRHEmplVacacionesPend, AR.cAreaDescripcion, IsNull((Select sum(IsNull(nMonto,0)) Monto from RHPlanillaDetCon Where cPersCod = PE.cPersCod And (cPlanillaCod = ('E01') Or cPlanillaCod = ('E06') or cPlanillaCod = ('E08')) And cRRHHPeriodo Like '201012%' And cRHConceptoCod like '1%' And cRHConceptoCod <> '112' And cRHConceptoCod <> '130' And cRHConceptoCod <> '194') ,0) nRHSueldoMonto"
If Mid(psPeriodo, 5, 2) >= "01" Then Sql = Sql & " ,Prov01 = IsNull((Select nProvision From MovVacaciones Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "01'" & " And cPersCod = PE.cPersCod), 0)"
If Mid(psPeriodo, 5, 2) >= "02" Then Sql = Sql & " ,Prov02 = IsNull((Select nProvision From MovVacaciones Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "02'" & " And cPersCod = PE.cPersCod), 0)"
If Mid(psPeriodo, 5, 2) >= "03" Then Sql = Sql & " ,Prov03 = IsNull((Select nProvision From MovVacaciones Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "03'" & " And cPersCod = PE.cPersCod), 0)"
If Mid(psPeriodo, 5, 2) >= "04" Then Sql = Sql & " ,Prov04 = IsNull((Select nProvision From MovVacaciones Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "04'" & " And cPersCod = PE.cPersCod), 0)"
If Mid(psPeriodo, 5, 2) >= "05" Then Sql = Sql & " ,Prov05 = IsNull((Select nProvision From MovVacaciones Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "05'" & " And cPersCod = PE.cPersCod), 0)"
If Mid(psPeriodo, 5, 2) >= "06" Then Sql = Sql & " ,Prov06 = IsNull((Select nProvision From MovVacaciones Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "06'" & " And cPersCod = PE.cPersCod), 0)"
If Mid(psPeriodo, 5, 2) >= "07" Then Sql = Sql & " ,Prov07 = IsNull((Select nProvision From MovVacaciones Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "07'" & " And cPersCod = PE.cPersCod), 0)"
If Mid(psPeriodo, 5, 2) >= "08" Then Sql = Sql & " ,Prov08 = IsNull((Select nProvision From MovVacaciones Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "08'" & " And cPersCod = PE.cPersCod), 0)"
If Mid(psPeriodo, 5, 2) >= "09" Then Sql = Sql & " ,Prov09 = IsNull((Select nProvision From MovVacaciones Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "09'" & " And cPersCod = PE.cPersCod), 0)"
If Mid(psPeriodo, 5, 2) >= "10" Then Sql = Sql & " ,Prov10 = IsNull((Select nProvision From MovVacaciones Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "10'" & " And cPersCod = PE.cPersCod), 0)"
If Mid(psPeriodo, 5, 2) >= "11" Then Sql = Sql & " ,Prov11 = IsNull((Select nProvision From MovVacaciones Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "11'" & " And cPersCod = PE.cPersCod), 0)"
If Mid(psPeriodo, 5, 2) >= "12" Then Sql = Sql & " ,Prov12 = IsNull((Select nProvision From MovVacaciones Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "12'" & " And cPersCod = PE.cPersCod), 0)"
Sql = Sql & " From RHPlanillaDet E Inner Join Persona PE On E.cPersCod = PE.cPersCod Inner Join RRHH EM On EM.cPersCod = E.cPersCod "
Sql = Sql & " Inner Join RHConcepto RHPD ON EM.cPersCod = RHPD.cPersCod And E.cRRHHPeriodo LIKE '" & psPeriodo & "%'"
Sql = Sql & " Inner Join RHEmpleado MC on PE.cPersCod = MC.cPersCod Inner Join Agencias A on EM.cAgenciaActual = A.cAgeCod"
Sql = Sql & " Inner Join Areas AR on AR.cAreaCod = EM.cAreaCodActual"
Sql = Sql & " GROUP BY cAgeDescripcion, PE.cPersCod, EM.cRHCod, PE.cPersNombre, EM.dIngreso, cAgenciaActual, nRHEmplVacacionesPend, AR.cAreaDescripcion"
Sql = Sql & " Order By cPersNombre"
    
If Co.AbreConexion Then
    Set CargarProvVacacionesConsol = Co.CargaRecordSet(Sql)
    Co.CierraConexion
End If
Set Rs = Nothing
Set Co = Nothing
End Function

Public Function CargarVacacionesGozadas(ByVal psPersCod As String) As ADODB.Recordset
Dim Sql As String
Dim Co As DConecta
Dim Rs As ADODB.Recordset
Set Co = New DConecta

Sql = "Select"
Sql = Sql & " RH.cPersCod, RH.DINGRESO, SubString(cRRHHPeriodo, 1, 6) cRRHHPeriodo,"
Sql = Sql & " Dias = (Select Sum(nMonto) From RHPlanillaDetCon RHPDCS Where cPlanillaCod = 'E06' And cRHConceptoCod = 413 And RH.cPersCod = cPersCod And SubString(RHPDCS.cRRHHPeriodo, 1, 6) Like SubString(PD.cRRHHPeriodo, 1, 6)), MontoVac = (Select Sum(nMonto) From RHPlanillaDetCon RHPDCS Where cPlanillaCod = 'E06' And cRHConceptoCod = 130 And RH.cPersCod = cPersCod And SubString(RHPDCS.cRRHHPeriodo, 1, 6) Like SubString(PD.cRRHHPeriodo, 1, 6))"
Sql = Sql & " From RHPlanillaDet PD Inner Join RRHH RH On RH.cPersCod = PD.cPersCod"
Sql = Sql & " Where cPlanillaCod = 'E06' And RH.cPersCod = '" & psPersCod & "'"
Sql = Sql & " Group By RH.cPersCod, RH.DINGRESO, RH.DINGRESO, SubString(cRRHHPeriodo, 1, 6) Order By SubString(cRRHHPeriodo, 1, 6) Asc"

If Co.AbreConexion Then
    Set CargarVacacionesGozadas = Co.CargaRecordSet(Sql)
    Co.CierraConexion
End If
Set Rs = Nothing
Set Co = Nothing
End Function
'***
