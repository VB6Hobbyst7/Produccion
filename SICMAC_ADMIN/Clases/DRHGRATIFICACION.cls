VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRHGratificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Function VerificaMes(ByVal psFechaMEs As String) As Boolean
Dim Co As DConecta
Dim rs As ADODB.Recordset
Dim sql As String

Set Co = New DConecta

sql = "Select * from MovGratificacion where cPeriodo='" & psFechaMEs & "' and cFlag is null "

Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
VerificaMes = False
If Not (rs.EOF And rs.BOF) Then
    VerificaMes = True
End If
Co.CierraConexion

Set Co = Nothing
Set rs = Nothing
End Function

'Public Function Get_Personal_Grati(ByVal pFecha As String, Optional psDiciembre As String = "00") As ADODB.Recordset
'Dim Co As DConecta
'Dim rs As ADODB.Recordset
'Dim sql As String
'
'Set rs = New ADODB.Recordset
'Set Co = New DConecta
'
'sql = " Declare @pdFecha datetime"
'sql = sql & " Declare @lsFecha varchar(8)"
'sql = sql & "     set @pdFecha = '" & Format(pFecha, "YYYYMMDD") & "'"
'sql = sql & " Set @lsFecha = Convert(varchar(8),@pdFecha,112)"
'
'
'sql = sql & " Select"
'sql = sql & " AAA.cPersCod,cPersNombre,cRHCod,nRHMesGratificacion,dIngreso,cAgenciaAsig,"
'sql = sql & " IsNull(  (select sum(IsNull(nRHConceptoValor,0)) Monto from rhconcepto where CPERSCOD =AAA.cPersCod and crhplanillacod = 'E01'and  cRHConceptoCod like '1%') ,0) nRHSueldoMonto,"
''Sql = Sql & " BBB.cPersCod,dRHSueldoFecha,nRHSueldoMonto,"
'sql = sql & " CCC.cPersCod , cRHContratoNro, nRHContratoTpo"
'sql = sql & " from ("
'sql = sql & "      select RH.cAgenciaAsig,RE.cPersCod,PE.cPersNombre, RH.cRHCod, RE.nRHEmplVacacionesPend, RH.dIngreso,nRHMesGratificacion"
'sql = sql & "      from rhempleado RE"
'sql = sql & "      Inner Join RRHH RH on RH.cPersCod = RE.cPersCod"
'sql = sql & "      Inner Join Persona PE on PE.cPersCod = RE.cPersCod"
''Sql = Sql & "      where not substring(convert(char(3),nRHEstado),1,1) in ('7','8') and substring(cRHCod,1,1)='E'"
'sql = sql & "     Inner Join"
'sql = sql & "     ("
'sql = sql & "         Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like  Left(@lsFecha,6) +  '%' And cPlanillaCod = 'E01' and cperscod not in (Select cPersCod from rrhh Where datediff(month,dcese,@pdFecha) = 0 and datediff(day,dcese,@pdFecha) <> 0 and crhcod like 'E%' and cperscod  in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%'  And cPlanillaCod = 'E01'))"
'sql = sql & "         Union"
'sql = sql & "         Select cPersCod from rrhh Where datediff(month,dingreso,@pdFecha) = 0 and dcese is null and crhcod like 'E%' and cperscod not in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%' And cPlanillaCod = 'E01')"
'sql = sql & "         Union"
'sql = sql & "         Select cPersCod from rrhh Where datediff(month,dingreso,@pdFecha) = 0 and crhcod like 'E%' and datediff(month,dcese,@pdFecha) = 0 and datediff(day,dcese,@pdFecha) = 0 and cperscod not in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%'  And cPlanillaCod = 'E01')"
'sql = sql & "         Union"
'sql = sql & "         Select cPersCod from rrhh Where datediff(month,dcese,@pdFecha) = 0 and datediff(day,dcese,@pdFecha) = 0 and crhcod like 'E%' and cperscod  in (Select cPersCod from RHPlanillaDet Where cRRHHPeriodo Like Left(@lsFecha,6) +  '%' And cPlanillaCod = 'E01')"
'sql = sql & "     ) E on E.cPersCod = PE.cPersCod"
'sql = sql & "  ) AAA"
''Sql = Sql & "  Inner Join ("
''Sql = Sql & "      Select cPersCod,dRHSueldoFecha,nRHSueldoMonto from RHSueldo RS where    dRHSueldoFecha = ("
''Sql = Sql & "      Select max(dRHSueldoFecha) from RHSueldo where  cPersCod = RS.cPersCod)"
''Sql = Sql & "  ) BBB on BBB.cPersCod = AAA.cPersCod"
'sql = sql & "  Inner Join ("
'sql = sql & "      Select cPersCod,cRHContratoNro,nRHContratoTpo,dFechaFin from RHContrato RE where    cRHContratoNro = ("
'sql = sql & "      Select max(cRHContratoNro) from RHContrato where  cPersCod = RE.cPersCod)"
'sql = sql & "  ) CCC on CCC.cPersCod = AAA.cPersCod"
'
'If psDiciembre = "12" Then
'    sql = sql & " where datediff(year,dIngreso,'" & pFecha & "') =0"
'Else
'    sql = sql & " where datediff(month,dIngreso,'" & pFecha & "') >=0"
'End If
'
'sql = sql & " Order by cRHCod "
'
'
'Co.AbreConexion
'Set rs = Co.CargaRecordSet(sql)
'Co.CierraConexion
'Set Get_Personal_Grati = rs
'
'Set rs = Nothing
'Set Co = Nothing
'
'
'
'End Function

Public Function Get_Personal_Grati(ByVal pFecha As String, Optional psDiciembre As String = "00") As ADODB.Recordset
Dim Co As DConecta
Dim rs As ADODB.Recordset
Dim sql As String

Set rs = New ADODB.Recordset
Set Co = New DConecta

'Sql = " Select"
'Sql = Sql & " AAA.cPersCod,cPersNombre,cRHCod,nRHMesGratificacion,dIngreso,"
'Sql = Sql & " IsNull(  (select sum(IsNull(nRHConceptoValor,0)) Monto from rhconcepto where CPERSCOD =AAA.cPersCod and crhplanillacod = 'E01'and  cRHConceptoCod like '1%') ,0) nRHSueldoMonto,"
''Sql = Sql & " BBB.cPersCod,dRHSueldoFecha,nRHSueldoMonto,"
'Sql = Sql & " CCC.cPersCod , cRHContratoNro, nRHContratoTpo"
'Sql = Sql & " from ("
'Sql = Sql & "      select RE.cPersCod,PE.cPersNombre, RH.cRHCod, RE.nRHEmplVacacionesPend, RH.dIngreso,nRHMesGratificacion"
'Sql = Sql & "      from rhempleado RE"
'Sql = Sql & "      Inner Join RRHH RH on RH.cPersCod = RE.cPersCod"
'Sql = Sql & "      Inner Join Persona PE on PE.cPersCod = RE.cPersCod"
'Sql = Sql & "      where not substring(convert(char(3),nRHEstado),1,1) in ('7','8') and substring(cRHCod,1,1)='E'"
'Sql = Sql & "  ) AAA"
''Sql = Sql & "  Inner Join ("
''Sql = Sql & "      Select cPersCod,dRHSueldoFecha,nRHSueldoMonto from RHSueldo RS where    dRHSueldoFecha = ("
''Sql = Sql & "      Select max(dRHSueldoFecha) from RHSueldo where  cPersCod = RS.cPersCod)"
''Sql = Sql & "  ) BBB on BBB.cPersCod = AAA.cPersCod"
'Sql = Sql & "  Inner Join ("
'Sql = Sql & "      Select cPersCod,cRHContratoNro,nRHContratoTpo,dFechaFin from RHContrato RE where    cRHContratoNro = ("
'Sql = Sql & "      Select max(cRHContratoNro) from RHContrato where  cPersCod = RE.cPersCod)"
'Sql = Sql & "  ) CCC on CCC.cPersCod = AAA.cPersCod"
'
'If psDiciembre = "12" Then
'    Sql = Sql & " where datediff(year,dIngreso,'" & pFecha & "') =0"
'Else
'    Sql = Sql & " where datediff(month,dIngreso,'" & pFecha & "') >=0"
'End If
'
'Sql = Sql & " Order by cRHCod "

'MAVM 20110715 ***
sql = " Select"
sql = sql & " AAA.cPersCod,cPersNombre,cRHCod,nRHMesGratificacion,dIngreso,"
sql = sql & " IsNull((Select sum(IsNull(nMonto,0)) Monto from RHPlanillaDetCon Where cPersCod = AAA.cPersCod And (cPlanillaCod = ('E01') Or cPlanillaCod = ('E06') or cPlanillaCod = ('E08')) And cRRHHPeriodo Like '" & Mid(Format(CDate(pFecha), gsFormatoMovFecha), 1, 6) & "%' And cRHConceptoCod like '1%' And cRHConceptoCod <> '112' And cRHConceptoCod <> '130') ,0) nRHSueldoMonto,"
sql = sql & " Round(ISNULL(Case When SubString(Convert(varchar (10), dIngreso, 112), 1, 4)+ SubString(Convert(varchar (10), dIngreso, 112), 5, 2) = '" & Mid(Format(CDate(pFecha), gsFormatoMovFecha), 1, 6) & "'"
sql = sql & " Then Round(Round(nRHSueldoMonto / 6, 2) * Round(Cast(Datediff(d, dIngreso, '" & Format(CDate(pFecha), gcFormatoFecha) & "') as Float )/30, 2, 1), 2)"
sql = sql & " Else Convert(Decimal (20, 2),(nRHSueldoMonto / 6))End , 0), 2) nRHGratiMonto,"
sql = sql & " Cast((Round(ISNULL(Case When SubString(Convert(varchar (10), dIngreso, 112), 1, 4)+ SubString(Convert(varchar (10), dIngreso, 112), 5, 2) = '" & Mid(Format(CDate(pFecha), gsFormatoMovFecha), 1, 6) & "'"
sql = sql & " Then Round(Round(nRHSueldoMonto / 6, 2) * Round(Cast(Datediff(d, dIngreso, '" & Format(CDate(pFecha), gcFormatoFecha) & "') as Decimal (20, 2) )/30, 2, 1), 2)"
sql = sql & " Else Convert(Decimal (20, 2),(nRHSueldoMonto / 6))End , 0), 2) * 0.09)As Decimal (20, 2)) nRHEsSaludMonto"
sql = sql & " from ("
'Sql = Sql & "      select RE.cPersCod,PE.cPersNombre, RH.cRHCod, RE.nRHEmplVacacionesPend, RH.dIngreso,nRHMesGratificacion"
'Sql = Sql & "      from rhempleado RE"
'Sql = Sql & "      Inner Join RRHH RH on RH.cPersCod = RE.cPersCod"
'Sql = Sql & "      Inner Join Persona PE on PE.cPersCod = RE.cPersCod"
'Sql = Sql & "      where not substring(convert(char(3),nRHEstado),1,1) in ('7','8') and substring(cRHCod,1,1)='E'"
'Sql = Sql & "  ) AAA"
'Sql = Sql & "  Inner Join ("
'Sql = Sql & "      Select cPersCod,cRHContratoNro,nRHContratoTpo,dFechaFin from RHContrato RE where    cRHContratoNro = ("
'Sql = Sql & "      Select max(cRHContratoNro) from RHContrato where  cPersCod = RE.cPersCod)"
'Sql = Sql & "  ) CCC on CCC.cPersCod = AAA.cPersCod"

sql = sql & "      Select PE.cPersCod, EM.cRHCod, PE.cPersNombre, EM.dIngreso, Round(Isnull(nRHMesGratificacion, 0),4) nRHMesGratificacion, IsNull((Select sum(IsNull(nMonto,0)) Monto from RHPlanillaDetCon Where cPersCod = PE.cPersCod And (cPlanillaCod = ('E01') Or cPlanillaCod = ('E06') or cPlanillaCod = ('E08')) And cRRHHPeriodo Like '" & Mid(Format(CDate(pFecha), gsFormatoMovFecha), 1, 6) & "%' And cRHConceptoCod like '1%' And cRHConceptoCod <> '112' And cRHConceptoCod <> '130') ,0) nRHSueldoMonto, cAgenciaActual "
sql = sql & "      From RHPlanillaDet E Inner Join Persona PE On E.cPersCod = PE.cPersCod"
sql = sql & "      Inner Join RRHH EM On EM.cPersCod = E.cPersCod"
sql = sql & "      Inner Join RHConcepto RHPD ON EM.CPERSCOD = RHPD.CPERSCOD"
sql = sql & "      And E.cRRHHPeriodo LIKE '" & Mid(Format(CDate(pFecha), gsFormatoMovFecha), 1, 6) & "%'"
'Sql = Sql & "      And E.cPlanillaCod = 'E01'"
sql = sql & "      And E.cPlanillaCod in ('E01','E08')"
sql = sql & "      Inner Join RHEmpleado MC on PE.cPersCod = MC.cPersCod"
sql = sql & "      GROUP BY PE.CPERSCOD, EM.CRHCOD, PE.CPERSNOMBRE, EM.DINGRESO, cAgenciaActual, nRHMesGratificacion"
sql = sql & "  ) AAA"


If psDiciembre = "12" Then
    sql = sql & " where datediff(year,dIngreso,'" & pFecha & "') =0"
End If

sql = sql & " Order by cPersNombre "

Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
Co.CierraConexion
Set Get_Personal_Grati = rs

Set rs = Nothing
Set Co = Nothing



End Function

Public Function MigraPeriodoGrati(ByVal psRHEmp As String, ByVal psUser As String, ByVal pdFecSis As String, _
ByVal psPeriodo As String, ByVal nMes As Integer, ByVal nCantMes As Integer) As Boolean
'cEstado = 2,  Migracion de la gratificacion del Primer periodo (Enero-Junio)
'cEstado = 3,  Migracion de la gratificacion del Segundo periodo (Julio -Diciembre)
Dim Sql1 As String
Dim Sql2 As String
Dim Co As DConecta
Dim cMes As String
Set Co = New DConecta

If nMes = 7 Then cMes = "2"
If nMes = 1 Then cMes = "3"

Sql1 = "Insert MovGratificacion "
Sql1 = Sql1 & "(cRHCod,cUser,dFecReg,cPeriodo,cEstado,nmes) values "
Sql1 = Sql1 & "('" & psRHEmp & "','" & psUser & "','" & pdFecSis & "','" & psPeriodo & "','" & cMes & "'," & nCantMes & " )"

Sql2 = " Update RHEmpleado set nRHMesGratificacion =0"
Sql2 = Sql2 & " From RHEmpleado RE"
Sql2 = Sql2 & " Inner Join RRHH RH on RH.cPersCod = RE.cPersCod"
Sql2 = Sql2 & " Where RH.cRHCod = '" & psRHEmp & "'"

On Error GoTo GratiErr
Co.AbreConexion
Co.BeginTrans
Co.Ejecutar (Sql1)
Co.Ejecutar (Sql2)
Co.CommitTrans
Co.CierraConexion

Set Co = Nothing
MigraPeriodoGrati = True
Exit Function

GratiErr:
    Co.RollBackTrans
    MigraPeriodoGrati = False
    Set Co = Nothing
End Function
Public Function VerificaMigracionGrati(ByVal psRHEmp As String, _
ByVal psPeriodo As String, ByVal nMes As Integer) As Boolean
Dim sql As String
Dim Co As DConecta
Dim rs As ADODB.Recordset
Dim Mes As String
Dim cMes As String
cMes = "1"
If nMes = 7 Then cMes = "2"
If nMes = 12 Then cMes = "3"

Set Co = New DConecta
sql = " Select * from MovGratificacion Where "
sql = sql & " cRHCod = '" & psRHEmp & "' and cPeriodo = '" & psPeriodo & "' and cFlag is Null and cEstado='" & cMes & "'"
VerificaMigracionGrati = False

Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
If Not (rs.EOF And rs.BOF) Then VerificaMigracionGrati = True
Co.CierraConexion

Set Co = Nothing
End Function

Public Function VerificaGratificacionMes(ByVal psRHEmp As String, ByVal psPeriodo As String) As Boolean
Dim sql As String
Dim rs As ADODB.Recordset
Dim Co As DConecta
Set Co = New DConecta
sql = " Select * from MovGratificacion Where "
sql = sql & " cRHCod = '" & psRHEmp & "' and cPeriodo = '" & psPeriodo & "' and cFlag is Null and cEstado='1'"
VerificaGratificacionMes = False

Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
If Not (rs.EOF And rs.BOF) Then VerificaGratificacionMes = True
Co.CierraConexion


Set rs = Nothing
Set Co = Nothing
End Function

'Public Function GrabaGratificacion(ByVal psRHEmp As String, ByVal psUser As String, _
'ByVal psPeriodo As String, ByVal pdFecSis As String, ByVal pnDias As Double) As Boolean
'Dim Sql1 As String
'Dim Sql2 As String
'Dim Sql3 As String
'Dim Co As DConecta
'' cEstado = 1, Incremento en uno la gratificacion
'
'Sql1 = ""
'Sql2 = "Insert MovGratificacion "
'Sql2 = Sql2 & "(cRHCod,cUser,dFecReg,cPeriodo,cEstado,nmes) values "
'Sql2 = Sql2 & "('" & psRHEmp & "','" & psUser & "','" & pdFecSis & "','" & psPeriodo & "','1',1)"
'
'Sql3 = "Update rhempleado set nRhMesGratificacion = nRhMesGratificacion + " & pnDias
'Sql3 = Sql3 & " From RHEmpleado RE"
'Sql3 = Sql3 & " Inner Join RRHH RH on RH.cPersCod = RE.cPersCod"
'Sql3 = Sql3 & " Where RH.cRHCod = '" & psRHEmp & "'"
'Set Co = New DConecta
'
'On Error GoTo GratiErr
'Co.AbreConexion
'Co.Ejecutar (Sql2)
'Co.Ejecutar (Sql3)
'Co.CierraConexion
'GrabaGratificacion = True
'Exit Function
'GratiErr:
'  Co.RollBackTrans
'  Set Co = Nothing
'  GrabaGratificacion = False
'  'MsgBox Err.Description, vbCritical, "AVISO"
'End Function

Public Function GrabaGratificacion(ByVal psRHEmp As String, ByVal psUser As String, _
ByVal psPeriodo As String, ByVal pdFecSis As String, ByVal pnDias As Double, Optional pnProvision As Currency) As Boolean
Dim Sql2 As String
Dim Sql3 As String
Dim Co As DConecta
' cEstado = 1, Incremento en uno la gratificacion

Sql2 = "Insert MovGratificacion "
Sql2 = Sql2 & "(cRHCod,cUser,dFecReg,cPeriodo,cEstado,nmes, nProvision) values "
Sql2 = Sql2 & "('" & psRHEmp & "','" & psUser & "','" & pdFecSis & "','" & psPeriodo & "','1',1, " & pnProvision & ")"

Sql3 = "Update RHEmpleado set nRhMesGratificacion = nRhMesGratificacion + " & pnDias
Sql3 = Sql3 & " From RHEmpleado RE"
Sql3 = Sql3 & " Inner Join RRHH RH on RH.cPersCod = RE.cPersCod"
Sql3 = Sql3 & " Where RH.cRHCod = '" & psRHEmp & "'"
Set Co = New DConecta

On Error GoTo GratiErr
Co.AbreConexion
Co.Ejecutar (Sql2)
Co.Ejecutar (Sql3)
Co.CierraConexion
GrabaGratificacion = True
Exit Function
GratiErr:
  Co.RollBackTrans
  Set Co = Nothing
  GrabaGratificacion = False
  'MsgBox Err.Description, vbCritical, "AVISO"
End Function

Public Function Get_Feriado(ByVal pFecha As String) As Date
Dim Co As DConecta
Dim sql As String
Dim rs As ADODB.Recordset
Set Co = New DConecta
sql = " Select top 1 * from DBComunes..Feriado "
sql = sql & " where datediff(month,dFeriado,'" & pFecha & "')=0 order by dFeriado"
Co.AbreConexion
Co.CargaRecordSet (sql)
Get_Feriado = rs!dFeriado
Co.CierraConexion

Set Co = Nothing
Set rs = Nothing
End Function

'MAVM 20110920 ***
Public Function CargarProvGratiConsol(ByVal psPeriodo As String) As ADODB.Recordset
Dim sql As String
Dim Co As DConecta
Dim rs As ADODB.Recordset
Set Co = New DConecta

sql = "Select"
sql = sql & " EM.cRHCod, PE.cPersNombre"
If (Mid(psPeriodo, 5, 2) = "01" Or Mid(psPeriodo, 5, 2) = "02" Or Mid(psPeriodo, 5, 2) = "03" Or Mid(psPeriodo, 5, 2) = "04" Or Mid(psPeriodo, 5, 2) = "05" Or Mid(psPeriodo, 5, 2) = "06") Then
    sql = sql & " , Prov01 = IsNull((Select nProvision From MovGratificacion Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "01'" & " And cRHCod = EM.cRHCod), 0)"
    sql = sql & " , Prov02 = IsNull((Select nProvision From MovGratificacion Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "02'" & " And cRHCod = EM.cRHCod), 0)"
    sql = sql & " , Prov03 = IsNull((Select nProvision From MovGratificacion Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "03'" & " And cRHCod = EM.cRHCod), 0)"
    sql = sql & " , Prov04 = IsNull((Select nProvision From MovGratificacion Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "04'" & " And cRHCod = EM.cRHCod), 0)"
    sql = sql & " , Prov05 = IsNull((Select nProvision From MovGratificacion Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "05'" & " And cRHCod = EM.cRHCod), 0)"
    sql = sql & " , Prov06 = IsNull((Select nProvision From MovGratificacion Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "06'" & " And cRHCod = EM.cRHCod), 0)"
End If

If (Mid(psPeriodo, 5, 2) = "07" Or Mid(psPeriodo, 5, 2) = "08" Or Mid(psPeriodo, 5, 2) = "09" Or Mid(psPeriodo, 5, 2) = "10" Or Mid(psPeriodo, 5, 2) = "11" Or Mid(psPeriodo, 5, 2) = "12") Then
    sql = sql & " , Prov07 = IsNull((Select nProvision From MovGratificacion Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "07'" & " And cRHCod = EM.cRHCod), 0)"
    sql = sql & " , Prov08 = IsNull((Select nProvision From MovGratificacion Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "08'" & " And cRHCod = EM.cRHCod), 0)"
    sql = sql & " , Prov09 = IsNull((Select nProvision From MovGratificacion Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "09'" & " And cRHCod = EM.cRHCod), 0)"
    sql = sql & " , Prov10 = IsNull((Select nProvision From MovGratificacion Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "10'" & " And cRHCod = EM.cRHCod), 0)"
    sql = sql & " , Prov11 = IsNull((Select nProvision From MovGratificacion Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "11'" & " And cRHCod = EM.cRHCod), 0)"
    sql = sql & " , Prov12 = IsNull((Select nProvision From MovGratificacion Where cPeriodo = '" & Mid(psPeriodo, 1, 4) & "12'" & " And cRHCod = EM.cRHCod), 0)"
End If

sql = sql & " From RHPlanillaDet E Inner Join Persona PE On E.cPersCod = PE.cPersCod Inner Join RRHH EM On EM.cPersCod = E.cPersCod "
sql = sql & " Inner Join RHConcepto RHPD ON EM.cPersCod = RHPD.cPersCod"

If (Mid(psPeriodo, 5, 2) = "01" Or Mid(psPeriodo, 5, 2) = "02" Or Mid(psPeriodo, 5, 2) = "03" Or Mid(psPeriodo, 5, 2) = "04" Or Mid(psPeriodo, 5, 2) = "05" Or Mid(psPeriodo, 5, 2) = "06") Then
    sql = sql & " And SubString(E.cRRHHPeriodo, 1, 6) Between '" & Mid(psPeriodo, 1, 4) & "01'" & "And '" & psPeriodo & "'"
End If
If (Mid(psPeriodo, 5, 2) = "07" Or Mid(psPeriodo, 5, 2) = "08" Or Mid(psPeriodo, 5, 2) = "09" Or Mid(psPeriodo, 5, 2) = "10" Or Mid(psPeriodo, 5, 2) = "11" Or Mid(psPeriodo, 5, 2) = "12") Then
    sql = sql & " And SubString(E.cRRHHPeriodo, 1, 6) Between '" & Mid(psPeriodo, 1, 4) & "01'" & "And '" & psPeriodo & "'"
End If

sql = sql & " And E.cPlanillaCod in ('E01') Inner Join RHEmpleado MC on PE.cPersCod = MC.cPersCod Inner Join Agencias A on EM.cAgenciaActual = A.cAgeCod"
sql = sql & " Inner Join Areas AR on AR.cAreaCod = EM.cAreaCodActual Where Not PE.cPersCod in ('1090101173286','1090101173294')"
sql = sql & " GROUP BY cAgeDescripcion, PE.cPersCod, EM.cRHCod, PE.cPersNombre, EM.dIngreso, cAgenciaActual, nRHEmplVacacionesPend, AR.cAreaDescripcion"
sql = sql & " Order By cPersNombre"
    
If Co.AbreConexion Then
    Set CargarProvGratiConsol = Co.CargaRecordSet(sql)
    Co.CierraConexion
End If
Set rs = Nothing
Set Co = Nothing
End Function
'MAVM 20110920 ***
