VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRiesgos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Function TiposGaratiaRiesgos() As ADODB.Recordset
Dim sql As String
Dim Co As DConecta
Set Co = New DConecta

sql = "select NCONSVALOR, CCONSDESCRIPCION"
sql = sql & " from constante WHERE NCONSCOD = 9060 AND NCONSVALOR != 9060"
sql = sql & " ORDER BY NCONSVALOR"

Co.AbreConexion
Set TiposGaratiaRiesgos = Co.CargaRecordSet(sql)
Co.CierraConexion
Set Co = Nothing
End Function

Function DiaCierreCalifMes(ByVal psServer As String) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim Co As DConecta

Set Co = New DConecta

Co.AbreConexion
sql = "select nConsSisValor from constsistema where nConsSisCod = 160"
Set rs = Co.CargaRecordSet(sql)
Co.CierraConexion
DiaCierreCalifMes = "-1"
If Not (rs.EOF And rs.BOF) Then
    DiaCierreCalifMes = Format(rs!nConssisValor, "DD/MM/YYYY")
Else
    MsgBox "No exite el codigo 155, avise a Sistemas", vbInformation, "AVISO"
End If
Set rs = Nothing
Set Co = Nothing
End Function
 
Function FechaCalificacion(ByVal psServer As String, ByVal Fecha As String, ByVal PersCod As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim Co As DConecta

sql = "select dfecha, cCtaCod, nDiasAtraso,"
sql = sql & " isnull(cCalGen,'-') Calgen,  isnull(cCalSistF,'-') CalSis,"
sql = sql & " isnull(cCalCMAC,'-') CMAC, isnull(cCalEval,'-') CalHist"
sql = sql & " From " & psServer & "ColocCalifProvTotal"
sql = sql & " where datediff(month,dfecha,'" & Fecha & "')=0"
sql = sql & " and cPersCod =  '" & PersCod & "'"

Set Co = New DConecta
Co.AbreConexion
Set FechaCalificacion = Co.CargaRecordSet(sql)
Co.CierraConexion

Set Co = Nothing
Set rs = Nothing

End Function

Function GetCreditodMesReclasificados(ByVal psServer As String, ByVal nTipoCambio As Double, ByVal psServerRCC As String, Optional ByVal pbComMes As Boolean = False) As ADODB.Recordset
Dim sql As String
Dim Co As DConecta

If pbComMes = False Then
    sql = " Select cPersCod, cPersNombre,"
    sql = sql & " Isnull(Sum(case when substring(Rd.cod_cuenta,3,1) = '2' And RD.cod_cuenta like '14_[1456]%' then Val_saldo / " & nTipoCambio & " End),0) nEndeudaDol"
    sql = sql & " From"
    sql = sql & " ("
    sql = sql & "     Select cPersNombre,CP.cPersCod, case when DNI =  null Or DNI ='' then Ruc else DNI end Doc from"
    sql = sql & "     (Select distinct C.cPersCod From " & psServer & "ColocCalifProvTotal C where nPrdEstado in (2020,2021,2022,2030,2031,2032) and substring(cCtaCod,6,1)='2') CP"
    sql = sql & "     Inner Join"
    sql = sql & "     (SELECT cPersCod , cPersNombre FROM PERSONA) P on P.cPersCod = Cp.cPersCod"
    sql = sql & "     Left Join"
    sql = sql & "      (Select cPersCod,cPersIDnro DNI from PersID where cPersIdTpo='1') PDNI on PDNI.cPersCod = CP.cPersCod"
    sql = sql & "     Left Join"
    sql = sql & "      (Select cPersCod,cPersIDnro RUC from PersID where cPersIdTpo='2') PRUC on PRuc.cPersCod = CP.cPersCod"
    sql = sql & " ) Per"
    sql = sql & " Inner Join " & psServerRCC & "RccTotal R on LTRIM(RTRIM(R.Cod_Doc_Id)) = LTRIM(RTRIM(Per.Doc))"
    sql = sql & " Inner Join " & psServerRCC & "RccTotalDet RD on LTRIM(RTRIM(R.Cod_sbs)) = LTRIM(RTRIM(RD.Cod_sbs))"
    sql = sql & " WHERE ( LTRIM(RTRIM(Per.doc)) <>'' or Per.doc <>null ) And"
    sql = sql & " RD.cod_cuenta like '14_[1456]%' and nEndeudaDol > 30000 "
    sql = sql & " Group by cPersCod,cPersNombre"
Else
    
    sql = " Select cPersCod, cPersNombre,"
    sql = sql & " Isnull(Sum(case when substring(Rd.cod_cuenta,3,1) = '2' And RD.cod_cuenta like '14_[1456]%' then Val_saldo / " & nTipoCambio & " End),0) nEndeudaDol"
    sql = sql & " From"
    sql = sql & " ("
    sql = sql & "     Select cPersNombre,CP.cPersCod, case when DNI =  null Or DNI ='' then Ruc else DNI end Doc from"
    sql = sql & "     (Select distinct C.cPersCod From " & psServer & "ColocCalifProvTotal C where nPrdEstado in (2020,2021,2022,2030,2031,2032) and substring(cCtaCod,6,1)='2') CP"
    sql = sql & "     Inner Join"
    sql = sql & "     (SELECT cPersCod , cPersNombre FROM PERSONA) P on P.cPersCod = Cp.cPersCod"
    sql = sql & "     Left Join"
    sql = sql & "      (Select cPersCod,cPersIDnro DNI from PersID where cPersIdTpo='1') PDNI on PDNI.cPersCod = CP.cPersCod"
    sql = sql & "     Left Join"
    sql = sql & "      (Select cPersCod,cPersIDnro RUC from PersID where cPersIdTpo='2') PRUC on PRuc.cPersCod = CP.cPersCod"
    sql = sql & " ) Per"
    sql = sql & " Inner Join " & psServerRCC & "RccTotal R on LTRIM(RTRIM(R.Cod_Doc_Id)) = LTRIM(RTRIM(Per.Doc))"
    sql = sql & " Inner Join " & psServerRCC & "RccTotalDet RD on LTRIM(RTRIM(R.Cod_sbs)) = LTRIM(RTRIM(RD.Cod_sbs))"
    sql = sql & " WHERE ( LTRIM(RTRIM(Per.doc)) <>'' or Per.doc <>null ) And"
    sql = sql & " RD.cod_cuenta like '14_[1456]%' "
    sql = sql & " Group by cPersCod,cPersNombre"
End If
Set Co = New DConecta
Co.AbreConexion
Set GetCreditodMesReclasificados = Co.CargaRecordSet(sql)
Co.CierraConexion
Set Co = Nothing
End Function

Function NuevoCredReclasificado(ByVal psFecSis As String, ByVal PersCod As String, ByVal psCodCMAC As String, _
ByVal psgColComercEmp As Integer, ByVal psMovNro As String, _
ByVal psCodAge As String, ByVal psCodUser As String) As Integer
Dim sql As String
Dim Co As DConecta
Dim Mo As DMov
Dim rs As ADODB.Recordset
Dim psNuevaCta As String
Dim psCuenta As String
Dim psMoneda As String
Dim oGen As DGeneral

Dim NroCred As Integer
Dim cMovNro As String

Dim Sql0 As String, SQL1 As String, Sql2 As String, Sql3 As String, Sql4 As String
Dim sql5 As String, Sql6 As String, Sql7 As String, sql8 As String, Sql9 As String
Dim Sql10 As String, Sql11 As String, Sql12 As String, Sql13 As String, Sql14 As String
Dim Sql15 As String, Sql16 As String
Set Co = New DConecta
Set oGen = New DGeneral
Set Mo = New DMov

Dim nMovNro As Long
If psgColComercEmp = "201" Then
    'sql = "Select cCtaCod, cPersCod from ProductoPersona where substring(cCtaCod,6,1)='1' "
    'sql = sql & " and nPrdPersRelac = 20 and cPersCod = '" & PersCod & "' and cEstado in (2020,2021,2022,2030,2031,2032)"
    
    sql = " Select  R.cCtaCod, R.cPersCod, P.nSaldo " _
        & " from    ProductoPersona R " _
        & "         JOIN PRODUCTO P ON P.cCtaCod = R.cCtaCod " _
        & " where   substring(R.cCtaCod,6,1)='1'  and r.nPrdPersRelac = 20 and R.cPersCod = '" & PersCod & "' " _
        & "         and nPrdEstado in (2020,2021,2022,2030,2031,2032) "
    
Else
    'sql = "Select cCtaCod, cPersCod from ProductoPersona where substring(cCtaCod,6,1)='2'"
    'sql = sql & " and nPrdPersRelac = 20 and cPersCod = '" & PersCod & "' and cEstado in (2020,2021,2022,2030,2031,2032)"
    
    sql = " Select  R.cCtaCod, R.cPersCod,P.nSaldo " _
        & " from    ProductoPersona R " _
        & "         JOIN PRODUCTO P ON P.cCtaCod = R.cCtaCod " _
        & " where   substring(R.cCtaCod,6,1)='2'  and r.nPrdPersRelac = 20 and R.cPersCod = '" & PersCod & "' " _
        & "         and nPrdEstado in (2020,2021,2022,2030,2031,2032) "
End If

On Error GoTo ErrorReclasificaMesComercial

Co.AbreConexion
'Co.BeginTrans
Set rs = Co.CargaRecordSet(sql)
NroCred = 0
While Not rs.EOF
    NroCred = NroCred + 1
    cMovNro = Mo.GeneraMovNro(psFecSis, psCodAge, psCodUser)
    nMovNro = Mo.InsertaMov(cMovNro, 172101, IIf(psgColComercEmp = "201", "Reclasifica Comercial a Mes", "Reclasificacion Mes Comercial"))
    nMovNro = Mo.GetnMovNro(cMovNro)
    psCuenta = rs!cCtaCod
    psMoneda = Mid(rs!cCtaCod, 9, 1)
    psNuevaCta = psCodCMAC & oGen.GeneraNuevaCuenta(Mid(rs!cCtaCod, 4, 2), psgColComercEmp, Mid(rs!cCtaCod, 9, 1))
        
    Sql15 = "INSERT INTO Producto(cCtaCod,nTasaInteres, nSaldo,nPrdEstado,dPrdEstado,nTransacc)"
    Sql15 = Sql15 & " Select  '" & psNuevaCta & "' cCtaCod,nTasaInteres, nSaldo,nPrdEstado,dPrdEstado,nTransacc"
    Sql15 = Sql15 & " From producto where cCtaCod = '" & psCuenta & "'"
    Co.Ejecutar (Sql15)
    
    Sql14 = "Insert ColocReclasificaMes"
    Sql14 = Sql14 & " (dFecha,cCtaCod,cCtaCodAnt,cCodope,nMovActualiza, nSaldoCap) values"
    Sql14 = Sql14 & "('" & Format(psFecSis, "mm/dd/yyyy") & "','" & psNuevaCta & "','" & psCuenta & "',"
    Sql14 = Sql14 & "'172101'," & nMovNro & "," & rs!nSaldo & ")"
    Co.Ejecutar (Sql14)
    
    sql = " Update ProductoPersona set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Co.Ejecutar (sql)
    
    SQL1 = " insert into colocaciones " _
        & " select '" & psNuevaCta & "', nPlazo, dVenc, nMontoCol, nColocCalendCod, cUltimaActualizacion, cLineaCred, dVigencia " _
        & " from colocaciones where cCtaCod ='" & psCuenta & "'"
    Co.Ejecutar (SQL1)
    
    Sql2 = " insert into colocaccred" _
        & " select  '" & psNuevaCta & "', nDiasAtraso, nDiasAtrasoAcum, nColocCondicion, nColocCondicionProd, nColocDestino, cProtesto, bCargoAuto, cMetLiquidacion, bRefCapInt, nNroProxCuota, nIntPend, " _
        & "         nExoPenalidad, nColocCalendCod,nCalendDinamico,nTipoDesembolso,nNroCalen,nNroProxDesemb,cPersCod,nFondoMoneda,cNumFte,dPersEval, bMiVivienda, bCuotaCom, nNroCalPar, nCalPago, bPrepago, " _
        & "         nCalendDinamTipo , nColocCondicion2, cCU, cRFA, IdCampana " _
        & "  from colocaccred where cCtaCod ='" & psCuenta & "'"

    Co.Ejecutar (Sql2)
    
    sql = " Update ColocCalificacionAnalista set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Co.Ejecutar (sql)
    
    sql = " Update ColocCalendario set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Co.Ejecutar (sql)
    
    sql = " Update ColocCalendDet set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Co.Ejecutar (sql)
    
    sql = " Update ColocCredCredVig set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Co.Ejecutar (sql)
    
    sql = " Update ColocGarantia set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Co.Ejecutar (sql)
    
    sql = " Update ColocGarantRec set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Co.Ejecutar (sql)
    
    sql = " Update ColocacEstado set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Co.Ejecutar (sql)
    
    sql = " Update ColocFteIngreso set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Co.Ejecutar (sql)
    
    sql = " Update movCol set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Co.Ejecutar (sql)
    
    sql = " Update MovColDet set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Co.Ejecutar (sql)
    
    sql = " delete colocaccred where cCtaCod ='" & psCuenta & "' "
    Co.Ejecutar (sql)
    sql = " delete colocaciones where cCtaCod ='" & psCuenta & "'"
    Co.Ejecutar (sql)
    sql = " delete from producto where cCtaCod ='" & psCuenta & "'"
    Co.Ejecutar (sql)
    
    rs.MoveNext
Wend
'Co.CommitTrans
Co.CierraConexion
'Codigo de reclasificacion Mes Comercial 172101
    Set Co = Nothing
    Set oGen = Nothing
    Set Mo = Nothing
    Set rs = Nothing
    NuevoCredReclasificado = NroCred
    
Exit Function
ErrorReclasificaMesComercial:
   ' Co.RollbackTrans
    Set Co = Nothing
    Set oGen = Nothing
    Set Mo = Nothing
    Set rs = Nothing
    NuevoCredReclasificado = -1
Call RaiseError(MyUnhandledError, "DRiesgo:NuevoCredReclasificado Method")
End Function

