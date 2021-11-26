VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMColocEval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'* Modulo de Colocaciones Pignoraticio
'* Clase con todas las Actualizaciones a BD
Option Explicit
Dim csConexion As String
Dim csNegocio As String
Dim csCentralPer As String
Dim csCentralCom As String
Dim csCentralImg As String
Dim csAdminist As String

Dim coConex As COMConecta.DCOMConecta

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
    
Set coConex = New COMConecta.DCOMConecta
If coConex.AbreConexion(csConexion) = False Then
    Err.Raise Err.Number, "Abrir Conexion", Err.Description
End If

End Sub

Private Sub Class_Terminate()
    coConex.CierraConexion
    Set coConex = Nothing
End Sub

Public Sub dInsertColocEvalCalif(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        ByVal psEvalCalif As String, ByVal psUltimaAct As String, ByVal psEvalObs As String, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT ColocEvalCalif (cPersCod, nEvalTipo, cEvalCalif, cUltimaAct, cEvalObs ) " _
        & "VALUES ('" & psPersCod & "'," & pnEvalTipo & ",'" & psEvalCalif & "','" _
        & psUltimaAct & "','" & psEvalObs & "'  )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub


Public Sub dInsertColocEvalCalifDetalle(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        ByVal psCtaCod As String, ByVal pdEval As Date, _
        ByVal psEvalEstado As String, ByVal psEvalCalifDet As String, _
        ByVal pnSaldoCap As Double, ByVal pnDiasAtraso As Integer, _
        ByVal psUltimaAct As String, ByVal psEvalObsDet As String, _
        Optional ByVal pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "INSERT ColocEvalCalifDetalle (cPersCod, nEvalTipo, cCtaCod, dEval, cEvalEstado, cEvalCalifDet, " _
        & "nSaldoCap, nDiasAtraso, cUltimaAct, cEvalObsDet ) " _
        & "VALUES ('" & psPersCod & "'," & pnEvalTipo & ",'" & psCtaCod & "','" & Format(pdEval, "mm/dd/yyyy") & "','" _
        & psEvalEstado & "','" & psEvalCalifDet & "'," & pnSaldoCap & "," & pnDiasAtraso & ",'" & psUltimaAct & "','" & psEvalObsDet & "'  )"
    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If
End Sub

Public Sub dUpdateColocEvalCalif(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        Optional ByVal psEvalCalif As String = "@", Optional ByVal psUltimaAct As String = "@", _
        Optional ByVal psEvalObs As String = "@", _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE ColocEvalCalif SET "

    If psEvalCalif <> "@" Then
        lsSQL = lsSQL & " cEvalCalif = '" & psEvalCalif & "',"
    End If
    If psUltimaAct <> "@" Then
        lsSQL = lsSQL & " cUltimaAct = '" & psUltimaAct & "',"
    End If
    If psEvalObs <> "@" Then
        lsSQL = lsSQL & " cEvalObs = '" & psEvalObs & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cPersCod ='" & psPersCod & "' AND nEvalTipo = " & pnEvalTipo & " "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dDeleteColocEvalCalif(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocEvalCalif " _
          & " WHERE cPersCod ='" & psPersCod & "' AND nEvalTipo = " & pnEvalTipo & " "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dUpdateColocEvalCalifDetalle(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        ByVal psCtaCod As String, ByVal pdEval As Date, Optional ByVal psEvalEstado As String, _
        Optional ByVal psEvalCalifDet As String = "@", Optional ByVal pnSaldoCap As Double = -999, _
        Optional ByVal pnDiasAtraso As Integer = -999, Optional ByVal psUltimaAct As String = "@", _
        Optional ByVal psEvalObsDet As String = "@", _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "UPDATE ColocEvalCalifDetalle SET "

    If psEvalEstado <> "@" Then
        lsSQL = lsSQL & " cEvalEstado = '" & psEvalEstado & "',"
    End If
    If psEvalCalifDet <> "@" Then
        lsSQL = lsSQL & " cEvalCalifDet = '" & psEvalCalifDet & "',"
    End If
    If pnSaldoCap <> -999 Then
         lsSQL = lsSQL & " nSaldoCap = " & pnSaldoCap & ","
    End If
    If pnDiasAtraso <> -999 Then
         lsSQL = lsSQL & " nDiasAtraso = " & pnDiasAtraso & ","
    End If
    If psUltimaAct <> "@" Then
        lsSQL = lsSQL & " cUltimaAct = '" & psUltimaAct & "',"
    End If
    If psEvalObsDet <> "@" Then
        lsSQL = lsSQL & " cEvalObsDet = '" & psEvalObsDet & "',"
    End If
    
    lsSQL = Left(lsSQL, Len(lsSQL) - 1)
    lsSQL = lsSQL & " WHERE cPersCod ='" & psPersCod & "' AND nEvalTipo = " & pnEvalTipo _
                  & " AND cCtaCod ='" & psCtaCod & "' AND dEval = '" & Format(pdEval, "yyyy/mm/dd") & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub

Public Sub dDeleteColocEvalCalifDetalle(ByVal psPersCod As String, ByVal pnEvalTipo As Integer, _
        ByVal psCtaCod As String, ByVal pdEval As Date, ByVal psEvalEstado As String, _
        Optional pbEjecBatch As Boolean = False)

Dim lsSQL As String
    lsSQL = "DELETE ColocEvalCalifDetalle  " _
          & " WHERE cPersCod ='" & psPersCod & "' AND nEvalTipo = " & pnEvalTipo _
                  & " AND cCtaCod ='" & psCtaCod & "' AND dEval = '" & Format(pdEval, "mm/dd/yyyy") & "' "

    If pbEjecBatch Then
        coConex.AdicionaCmdBatch lsSQL
    Else
        coConex.Ejecutar lsSQL
    End If

End Sub


'************************************************************
'***    DE ACA EN ADELANTE SON FUNCIONES DE LA CLASE DE DRiesgos
'************************************************************

Function TiposGaratiaRiesgos() As ADODB.Recordset
Dim SQL As String
'Dim Co As COMConecta.DCOMConecta
'Set Co = New COMConecta.DCOMConecta

SQL = "select NCONSVALOR, CCONSDESCRIPCION"
SQL = SQL & " from constante WHERE NCONSCOD = 9060 AND NCONSVALOR != 9060"
SQL = SQL & " ORDER BY NCONSVALOR"

'coConex.AbreConexion
Set TiposGaratiaRiesgos = coConex.CargaRecordSet(SQL)
'Co.CierraConexion
'Set Co = Nothing
End Function

Function DiaCierreCalifMes(ByVal psServer As String, _
                            Optional ByRef psMensaje As String = "") As String
Dim SQL As String
Dim rs As ADODB.Recordset
'Dim Co As COMConecta.DCOMConecta

'Set Co = New COMConecta.DCOMConecta

'Co.AbreConexion
SQL = "select nConsSisValor from constsistema where nConsSisCod = 160"
Set rs = coConex.CargaRecordSet(SQL)
'Co.CierraConexion
DiaCierreCalifMes = "-1"
If Not (rs.EOF And rs.BOF) Then
    DiaCierreCalifMes = Format(rs!nConsSisValor, "DD/MM/YYYY")
Else
    psMensaje = "No exite el codigo 155, avise a Sistemas"
End If
Set rs = Nothing
'Set Co = Nothing
End Function
 
Function FechaCalificacion(ByVal psServer As String, ByVal Fecha As String, ByVal PersCod As String) As ADODB.Recordset
Dim SQL As String
Dim rs As ADODB.Recordset
'Dim Co As COMConecta.DCOMConecta

SQL = "select dfecha, cCtaCod, nDiasAtraso,"
SQL = SQL & " isnull(cCalGen,'-') Calgen,  isnull(cCalSistF,'-') CalSis,"
SQL = SQL & " isnull(cCalCMAC,'-') CMAC, isnull(cCalEval,'-') CalHist"
SQL = SQL & " From " & psServer & "ColocCalifProvTotal"
SQL = SQL & " where datediff(month,dfecha,'" & Fecha & "')=0"
SQL = SQL & " and cPersCod =  '" & PersCod & "'"

'Set Co = New COMConecta.DCOMConecta
'co.AbreConexion
Set FechaCalificacion = coConex.CargaRecordSet(SQL)
'Co.CierraConexion

'Set Co = Nothing
Set rs = Nothing

End Function

Function GetCreditodMesReclasificados(ByVal psServer As String, ByVal nTipoCambio As Double, ByVal psServerRCC As String, Optional ByVal pbComMes As Boolean = False) As ADODB.Recordset
Dim SQL As String
'Dim Co As COMConecta.DCOMConecta

If pbComMes = False Then
    SQL = " Select cPersCod, cPersNombre,"
    SQL = SQL & " Isnull(Sum(case when substring(Rd.cod_cuenta,3,1) = '2' And RD.cod_cuenta like '14_[1456]%' then Val_saldo / " & nTipoCambio & " End),0) nEndeudaDol"
    SQL = SQL & " From"
    SQL = SQL & " ("
    SQL = SQL & "     Select cPersNombre,CP.cPersCod, case when DNI =  null Or DNI ='' then Ruc else DNI end Doc from"
    SQL = SQL & "     (Select distinct C.cPersCod From " & psServer & "ColocCalifProvTotal C where nPrdEstado in (2020,2021,2022,2030,2031,2032) and substring(cCtaCod,6,1)='2') CP"
    SQL = SQL & "     Inner Join"
    SQL = SQL & "     (SELECT cPersCod , cPersNombre FROM PERSONA) P on P.cPersCod = Cp.cPersCod"
    SQL = SQL & "     Left Join"
    SQL = SQL & "      (Select cPersCod,cPersIDnro DNI from PersID where cPersIdTpo='1') PDNI on PDNI.cPersCod = CP.cPersCod"
    SQL = SQL & "     Left Join"
    SQL = SQL & "      (Select cPersCod,cPersIDnro RUC from PersID where cPersIdTpo='2') PRUC on PRuc.cPersCod = CP.cPersCod"
    SQL = SQL & " ) Per"
    SQL = SQL & " Inner Join " & psServerRCC & "RccTotal R on LTRIM(RTRIM(R.Cod_Doc_Id)) = LTRIM(RTRIM(Per.Doc))"
    SQL = SQL & " Inner Join " & psServerRCC & "RccTotalDet RD on LTRIM(RTRIM(R.Cod_Edu)) = LTRIM(RTRIM(RD.Cod_Edu))"
    SQL = SQL & " WHERE ( LTRIM(RTRIM(Per.doc)) <>'' or Per.doc <>null ) And"
    SQL = SQL & " RD.cod_cuenta like '14_[1456]%' " 'and nEndeudaDol > 30000 "
    SQL = SQL & " Group by cPersCod,cPersNombre"
    SQL = SQL & " Having Isnull(Sum(case when substring(Rd.cod_cuenta,3,1) = '2' And RD.cod_cuenta like '14_[1456]%' then Val_saldo / " & nTipoCambio & " End),0)> 3000"
Else
    
    SQL = " Select cPersCod, cPersNombre,"
    SQL = SQL & " Isnull(Sum(case when substring(Rd.cod_cuenta,3,1) = '2' And RD.cod_cuenta like '14_[1456]%' then Val_saldo / " & nTipoCambio & " End),0) nEndeudaDol"
    SQL = SQL & " From"
    SQL = SQL & " ("
    SQL = SQL & "     Select cPersNombre,CP.cPersCod, case when DNI =  null Or DNI ='' then Ruc else DNI end Doc from"
    SQL = SQL & "     (Select distinct C.cPersCod From " & psServer & "ColocCalifProvTotal C where nPrdEstado in (2020,2021,2022,2030,2031,2032) and substring(cCtaCod,6,1)='2') CP"
    SQL = SQL & "     Inner Join"
    SQL = SQL & "     (SELECT cPersCod , cPersNombre FROM PERSONA) P on P.cPersCod = Cp.cPersCod"
    SQL = SQL & "     Left Join"
    SQL = SQL & "      (Select cPersCod,cPersIDnro DNI from PersID where cPersIdTpo='1') PDNI on PDNI.cPersCod = CP.cPersCod"
    SQL = SQL & "     Left Join"
    SQL = SQL & "      (Select cPersCod,cPersIDnro RUC from PersID where cPersIdTpo='2') PRUC on PRuc.cPersCod = CP.cPersCod"
    SQL = SQL & " ) Per"
    SQL = SQL & " Inner Join " & psServerRCC & "RccTotal R on LTRIM(RTRIM(R.Cod_Doc_Id)) = LTRIM(RTRIM(Per.Doc))"
    SQL = SQL & " Inner Join " & psServerRCC & "RccTotalDet RD on LTRIM(RTRIM(R.Cod_Edu)) = LTRIM(RTRIM(RD.Cod_Edu))"
    SQL = SQL & " WHERE ( LTRIM(RTRIM(Per.doc)) <>'' or Per.doc <>null ) And"
    SQL = SQL & " RD.cod_cuenta like '14_[1456]%' "
    SQL = SQL & " Group by cPersCod,cPersNombre"
End If
'Set Co = New COMConecta.DCOMConecta
'Co.AbreConexion
Set GetCreditodMesReclasificados = coConex.CargaRecordSet(SQL)
'Co.CierraConexion
'Set Co = Nothing
End Function

Function NuevoCredReclasificado(ByVal psFecSis As String, ByVal PersCod As String, ByVal psCodCmac As String, _
                            ByVal psgColComercEmp As Integer, ByVal psMovNro As String, _
                            ByVal psCodAge As String, ByVal psCodUser As String) As Integer
Dim SQL As String
'Dim Co As COMConecta.DCOMConecta
Dim Mo As COMDMov.DCOMMov
Dim rs As ADODB.Recordset
Dim psNuevaCta As String
Dim psCuenta As String
Dim psMoneda As String
Dim oGen As COMDConstSistema.DCOMGeneral

Dim NroCred As Integer
Dim cMovnro As String

Dim Sql0 As String, SQL1 As String, Sql2 As String, Sql3 As String, Sql4 As String
Dim sql5 As String, Sql6 As String, Sql7 As String, sql8 As String, Sql9 As String
Dim Sql10 As String, Sql11 As String, Sql12 As String, Sql13 As String, Sql14 As String
Dim Sql15 As String, Sql16 As String
'Set Co = New COMConecta.DCOMConecta
Set oGen = New COMDConstSistema.DCOMGeneral
Set Mo = New COMDMov.DCOMMov

Dim nMovNro As Long
If psgColComercEmp = "201" Then
    'sql = "Select cCtaCod, cPersCod from ProductoPersona where substring(cCtaCod,6,1)='1' "
    'sql = sql & " and nPrdPersRelac = 20 and cPersCod = '" & PersCod & "' and cEstado in (2020,2021,2022,2030,2031,2032)"
    
    SQL = " Select  R.cCtaCod, R.cPersCod, P.nSaldo " _
        & " from    ProductoPersona R " _
        & "         JOIN PRODUCTO P ON P.cCtaCod = R.cCtaCod " _
        & " where   substring(R.cCtaCod,6,1)='1'  and r.nPrdPersRelac = 20 and R.cPersCod = '" & PersCod & "' " _
        & "         and nPrdEstado in (2020,2021,2022,2030,2031,2032) "
    
Else
    'sql = "Select cCtaCod, cPersCod from ProductoPersona where substring(cCtaCod,6,1)='2'"
    'sql = sql & " and nPrdPersRelac = 20 and cPersCod = '" & PersCod & "' and cEstado in (2020,2021,2022,2030,2031,2032)"
    
    SQL = " Select  R.cCtaCod, R.cPersCod,P.nSaldo " _
        & " from    ProductoPersona R " _
        & "         JOIN PRODUCTO P ON P.cCtaCod = R.cCtaCod " _
        & " where   substring(R.cCtaCod,6,1)='2'  and r.nPrdPersRelac = 20 and R.cPersCod = '" & PersCod & "' " _
        & "         and nPrdEstado in (2020,2021,2022,2030,2031,2032) "
End If

On Error GoTo ErrorReclasificaMesComercial

'Co.AbreConexion
'Co.BeginTrans
Set rs = coConex.CargaRecordSet(SQL)
NroCred = 0
While Not rs.EOF
    NroCred = NroCred + 1
    cMovnro = Mo.GeneraMovNro(psFecSis, psCodAge, psCodUser)
    nMovNro = Mo.InsertaMov(cMovnro, 172101, IIf(psgColComercEmp = "201", "Reclasifica Comercial a Mes", "Reclasificacion Mes Comercial"))
    nMovNro = Mo.GetnMovNro(cMovnro)
    psCuenta = rs!cCTaCod
    psMoneda = Mid(rs!cCTaCod, 9, 1)
    psNuevaCta = psCodCmac & oGen.GeneraNuevaCuenta(Mid(rs!cCTaCod, 4, 2), psgColComercEmp, Mid(rs!cCTaCod, 9, 1))
        
    Sql15 = "INSERT INTO Producto(cCtaCod,nTasaInteres, nSaldo,nPrdEstado,dPrdEstado,nTransacc)"
    Sql15 = Sql15 & " Select  '" & psNuevaCta & "' cCtaCod,nTasaInteres, nSaldo,nPrdEstado,dPrdEstado,nTransacc"
    Sql15 = Sql15 & " From producto where cCtaCod = '" & psCuenta & "'"
    coConex.Ejecutar (Sql15)
    
    Sql14 = "Insert ColocReclasificaMes"
    Sql14 = Sql14 & " (dFecha,cCtaCod,cCtaCodAnt,cCodope,nMovActualiza, nSaldoCap) values"
    Sql14 = Sql14 & "('" & Format(psFecSis, "mm/dd/yyyy") & "','" & psNuevaCta & "','" & psCuenta & "',"
    Sql14 = Sql14 & "'172101'," & nMovNro & "," & rs!nSaldo & ")"
    coConex.Ejecutar (Sql14)
    
    SQL = " Update ProductoPersona set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    coConex.Ejecutar (SQL)
    
    SQL1 = " insert into colocaciones " _
        & " select '" & psNuevaCta & "', nPlazo, dVenc, nMontoCol, nColocCalendCod, cUltimaActualizacion, cLineaCred, dVigencia " _
        & " from colocaciones where cCtaCod ='" & psCuenta & "'"
    coConex.Ejecutar (SQL1)
    
    Sql2 = " insert into colocaccred" _
        & " select  '" & psNuevaCta & "', nDiasAtraso, nDiasAtrasoAcum, nColocCondicion, nColocCondicionProd, nColocDestino, cProtesto, bCargoAuto, cMetLiquidacion, bRefCapInt, nNroProxCuota, nIntPend, " _
        & "         nExoPenalidad, nColocCalendCod,nCalendDinamico,nTipoDesembolso,nNroCalen,nNroProxDesemb,cPersCod,nFondoMoneda,cNumFte,dPersEval, bMiVivienda, bCuotaCom, nNroCalPar, nCalPago, bPrepago, " _
        & "         nCalendDinamTipo , nColocCondicion2, cCU, cRFA, IdCampana " _
        & "  from colocaccred where cCtaCod ='" & psCuenta & "'"

    coConex.Ejecutar (Sql2)
    
    SQL = " Update ColocCalificacionAnalista set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    coConex.Ejecutar (SQL)
    
    SQL = " Update ColocCalendario set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    coConex.Ejecutar (SQL)
    
    SQL = " Update ColocCalendDet set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    coConex.Ejecutar (SQL)
    
    SQL = " Update ColocCredCredVig set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    coConex.Ejecutar (SQL)
    
    SQL = " Update ColocGarantia set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    coConex.Ejecutar (SQL)
    
    SQL = " Update ColocGarantRec set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    coConex.Ejecutar (SQL)
    
    SQL = " Update ColocacEstado set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    coConex.Ejecutar (SQL)
    
    SQL = " Update ColocFteIngreso set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    coConex.Ejecutar (SQL)
    
    SQL = " Update movCol set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    coConex.Ejecutar (SQL)
    
    SQL = " Update MovColDet set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    coConex.Ejecutar (SQL)
    
    SQL = " delete colocaccred where cCtaCod ='" & psCuenta & "' "
    coConex.Ejecutar (SQL)
    SQL = " delete colocaciones where cCtaCod ='" & psCuenta & "'"
    coConex.Ejecutar (SQL)
    SQL = " delete from producto where cCtaCod ='" & psCuenta & "'"
    coConex.Ejecutar (SQL)
    
    rs.MoveNext
Wend
'Co.CommitTrans
'Co.CierraConexion
'Codigo de reclasificacion Mes Comercial 172101
'    Set Co = Nothing
    Set oGen = Nothing
    Set Mo = Nothing
    Set rs = Nothing
    NuevoCredReclasificado = NroCred
    
Exit Function
ErrorReclasificaMesComercial:
   ' Co.RollbackTrans
'    Set Co = Nothing
    Set oGen = Nothing
    Set Mo = Nothing
    Set rs = Nothing
    NuevoCredReclasificado = -1
    Err.Raise Err.Number, "Nuevo Credito Refinanciado", Err.Description
End Function

'**************************************************
'**** DE ACA EN ADELANTE SON FUNCIONES DE LA CLASE DRCC
'**************************************************

Sub BorraTablaRCC(ByVal psServer As String)
Dim SQL1 As String
Dim Sql2 As String

'Dim Co As DConecta

'Set Co = New DConecta

SQL1 = "delete " & psServer & "rcctotal"
Sql2 = "delete " & psServer & "rcctotaldet"


On Error GoTo RCC
'    Co.AbreConexion
    coConex.Ejecutar (SQL1)
    coConex.Ejecutar (Sql2)
'    Co.CierraConexion
'    Set Co = Nothing
Exit Sub
RCC:
'    Set Co = Nothing
   Err.Raise Err.Number, "Borra Tabla RCC", Err.Description
End Sub

Sub InsertaRccHistorico(ByVal pcConexServerConsol As ADODB.Connection)
Dim SQL1 As String
Dim Sql2 As String
'Dim Co As DConecta

'Set Co = New DConecta


SQL1 = " Insert rcchistotal"
SQL1 = SQL1 & " Select * from rcctotal"

Sql2 = "Insert rcchistotaldet"
Sql2 = Sql2 & " Select * from rcctotaldet"

pcConexServerConsol.CommandTimeout = 12000
pcConexServerConsol.Execute SQL1
pcConexServerConsol.Execute Sql2

Exit Sub
RCC:
'   Set Co = Nothing
   Err.Raise Err.Number, "Inserta RCC Historico", Err.Description
End Sub


Function GetFecha(ByVal psServer As String, ByVal nTpo As Integer, _
                    Optional ByRef psMensaje As String = "") As String
'Definicion
'1 RccTotal
'2 RccTotalDet
'3 RccHisTotal
'4 RccHisTotalDet
Dim SQL As String ', psMensaje As String
'Dim Con As DConecta
Dim rs As ADODB.Recordset
Dim sFecha As String

Select Case nTpo
    Case 1
        SQL = "select max(Fec_Rep) Fecha from " & psServer & "RccTotal"
        psMensaje = " en RccTotal "
    Case 2
        SQL = "select max(dFecha) Fecha from " & psServer & "RccTotalDet"
        psMensaje = " en RccTotalDet "
    Case 3
        SQL = "select max(Fec_Rep) Fecha from " & psServer & "RccHisTotal "
        psMensaje = " en RccHisTotal "
    Case 4
        SQL = "select max(dFecha) Fecha from " & psServer & "RccHisTotalDet"
        psMensaje = " en RccHisTotalDet "
End Select

'Set Con = New DConecta

'con.AbreConexion
Set rs = coConex.CargaRecordSet(SQL)
If Not (rs.EOF And rs.BOF) Then
    sFecha = IIf(IsNull(rs!Fecha), "-1", rs!Fecha)
Else
    sFecha = "-1"
End If

'Con.CierraConexion

If sFecha = "-1" Then
    psMensaje = "No exiten datos " & psMensaje
End If

GetFecha = Format(sFecha, "DD/MM/YYYY")
Set rs = Nothing
'Set Co = Nothing
End Function

Sub InsertaRccCab(ByVal psServer As String, ByVal sFecCD As String)
Dim SQL As String
'Dim Co As DConecta

'Set Co = New DConecta

SQL = " Insert " & psServer & "rcctotal"

SQL = SQL & " Select"
SQL = SQL & " convert(datetime,substring(Col001,12,8)) Fec_Rep,"
SQL = SQL & " substring(Col001,2,10) Cod_Edu, null Tip_Det,"
SQL = SQL & " substring(Col001,20,1) Tip_Doc_Trib,"
SQL = SQL & " substring(Col001,21,11) Cod_Doc_Trib,"
SQL = SQL & " substring(Col001,32,1) Tip_Doc_Id,"
SQL = SQL & " substring(Col001,33,12) Cod_Doc_Id,"
SQL = SQL & " substring(Col001,45,1) TipPers,"
SQL = SQL & " substring(Col001,46,1) TipEmp,"
SQL = SQL & " convert(int,substring(Col001,47,3)) Can_ents,"
SQL = SQL & " isnull(convert(real,substring(Col001,50,5)/100),'') Calif_0,"
SQL = SQL & " isnull(convert(real,substring(Col001,55,5)/100),'') Calif_1,"
SQL = SQL & " isnull(convert(real,substring(Col001,60,5)/100),'') Calif_2,"
SQL = SQL & " isnull(convert(real,substring(Col001,65,5)/100),'') Calif_3,"
SQL = SQL & " isnull(convert(real,substring(Col001,70,5)/100),'') Calif_4,"
SQL = SQL & " case when len(rtrim(ltrim(substring(Col001,75,120))))=0 then '' else rtrim(ltrim(substring(Col001,75,120))) + ' ' end +"
SQL = SQL & " case when len(IsNull(rtrim(ltrim(substring(Col001,195,40))),'')) = 0 then '' else  + IsNull(rtrim(ltrim(substring(Col001,195,40))),'')  + ' ' end +"
SQL = SQL & " case when len(rtrim(ltrim(substring(Col001,235,40))))=0 then '' else rtrim(ltrim(substring(Col001,235,40))) + ' ' end +"
SQL = SQL & " case when len(rtrim(ltrim(substring(Col001,275,40))))=0 then '' else rtrim(ltrim(substring(Col001,275,40))) + ' ' end +"
SQL = SQL & " case when len(rtrim(ltrim(substring(Col001,315,40))))=0 then '' else rtrim(ltrim(substring(Col001,315,40))) end Nom_Deu"
SQL = SQL & " From " & psServer & "rcc" & sFecCD
SQL = SQL & " where substring(Col001,1,1)='1'"


On Error GoTo RccErr
'    Co.AbreConexion
    'Co.AdicionaCmdBatch (Sql)
    coConex.ConexionActiva.CommandTimeout = 12000
    coConex.Ejecutar (SQL)
    
'    Co.CierraConexion
'    Set Co = Nothing
Exit Sub

RccErr:
   Err.Raise Err.Number, "Inserta RCC Cab", Err.Description
'   Set Co = Nothing
End Sub

Sub InsertRCCDet(ByVal psServer As String, ByVal sFecCD As String)

Dim SQL As String
Dim sF As String
'Dim Co As DConecta

'Set Co = New DConecta
'Calcula ultimo dia del mes
sF = "01/" & Right(sFecCD, 2) & "/" & Left(sFecCD, 4)
sF = DateAdd("M", 1, sF)
sF = DateAdd("D", -1, sF)

SQL = "Insert " & psServer & "rcctotaldet"
SQL = SQL & "(dFecha,Cod_Edu,Cod_Emp,Tip_Credito, Cod_Cuenta,Condicion,Val_Saldo,Clasificacion)"

SQL = SQL & " Select"
SQL = SQL & " convert(datetime,'" & Format(sF, "YYYY/MM/DD") & "') dFecha,"
SQL = SQL & " substring(Col001,2,10) Cod_Edu,"
SQL = SQL & " substring(Col001,12,5) Cod_Emp,"
SQL = SQL & " substring(Col001,17,1) Tip_Credito,"
SQL = SQL & " isnull(substring(Col001,18,14),'') Cod_Cuenta,"
SQL = SQL & " case when isnumeric(substring(Col001,32,4)) = 1 then substring(Col001,32,4) else  null end Condicion,"
SQL = SQL & " convert(money,substring(Col001,36,15))/100 Val_Saldo,"
SQL = SQL & " substring(Col001,51,1) Clasificacion"
SQL = SQL & " From " & psServer & "rcc" & sFecCD
SQL = SQL & " where substring(Col001,1,1)='2'"


On Error GoTo RccErr
'    Co.AbreConexion
    coConex.ConexionActiva.CommandTimeout = 12000
    coConex.Ejecutar (SQL)
'    Co.CierraConexion
'    Set Co = Nothing
Exit Sub

RccErr:
   Err.Raise Err.Number, "Inserta RCC Det", Err.Description
'   Set Co = Nothing
End Sub

Sub InsertaDataColocCalifProvTotal(ByVal sServer As String, ByVal psFecData As String)
Dim SQL As String
'Dim Co As DConecta
'Set Co = New DConecta
SQL = " INSERT " & sServer & "ColocCalifProvTotal "
SQL = SQL & " SELECT '" & Format(psFecData, "YYYY/MM/DD") & "', * FROM ColocCalifProv"
'Co.AbreConexion
coConex.Ejecutar (SQL)
'Co.CierraConexion
'Set Co = Nothing
End Sub

Sub ActulizaFechaCierre(ByVal psFecData As String)
Dim SQL As String
'Dim Co As DConecta
'Set Co = New DConecta
'Co.AbreConexion
SQL = "UPDATE CONSTSISTEMA SET NCONSSISVALOR= '" & psFecData & "' where nConsSisCod = 160"
coConex.Ejecutar (SQL)
'Co.CierraConexion
'Set Co = Nothing
End Sub

Function VerificaDataMigradaFecha(ByVal psServer As String, ByVal psFecData As String, _
                                Optional ByRef psMensaje As String = "") As Long
Dim SQL As String
'Dim Co As DConecta
Dim rs As ADODB.Recordset

'Set Co = New DConecta

SQL = " select count(*) nro from " & psServer & "ColocCalifProvTotal "
SQL = SQL & " where datediff(day,dfecha,'" & Format(psFecData, "YYYY/MM/DD") & "')=0"

'Co.AbreConexion
Set rs = coConex.CargaRecordSet(SQL)
'Co.CierraConexion
VerificaDataMigradaFecha = rs!Nro
If rs!Nro > 0 Then
    psMensaje = "La Data ya fue Transferida"
    Exit Function
End If
'Set Co = Nothing
Set rs = Nothing
End Function
Function FechasRcc(ByVal psServer As String) As ADODB.Recordset
Dim SQL As String
'Dim Co As DConecta
Dim rs As ADODB.Recordset

'Set Co = New DConecta

SQL = "select distinct fec_Rep Fecha from " & psServer & "rcchistotal"
'Co.AbreConexion
Set rs = coConex.CargaRecordSet(SQL)
'Co.CierraConexion
Set FechasRcc = rs

'Set Co = Nothing
Set rs = Nothing
End Function

Function ServConsol(ByVal pnServer As Integer) As String
Dim SQL As String
Dim rs As ADODB.Recordset
'Dim Co As DConecta
SQL = "Select nConsSisValor serv from ConstSistema where nConsSisCod = " & pnServer
'Set Co = New DConecta
'Co.AbreConexion
Set rs = coConex.CargaRecordSet(SQL)
If Not (rs.EOF And rs.BOF) Then
    ServConsol = Trim(rs!Serv)
End If
'Co.CierraConexion
Set rs = Nothing
'Set Co = Nothing
End Function

Function NombreServerConsol() As String
Dim SQL As String
Dim rs As ADODB.Recordset
'Dim Co As DConecta
'Set Co = New DConecta
SQL = "SELECT Substring(nConsSisValor,1,LEN(nConsSisValor) - 5) SERVIDOR FROM ConstSistema WHERE nConsSisCod  = 43"
'Co.AbreConexion
Set rs = coConex.CargaRecordSet(SQL)
If Not (rs.EOF And rs.BOF) Then
    NombreServerConsol = Trim(rs!SERVIDOR)
End If
'Co.CierraConexion
Set rs = Nothing
'Set Co = Nothing
End Function

Function VerificaTabla(ByVal psServer As String, ByVal sFecCD As String) As Integer
Dim SQL As String
'Dim Co As DConecta
Dim rs As ADODB.Recordset

SQL = " select * from " & psServer & "sysobjects  where name = 'Rcc" & sFecCD & "'"
'Set Co = New DConecta
'Co.AbreConexion

Set rs = coConex.CargaRecordSet(SQL)
If Not (rs.EOF And rs.BOF) Then
    VerificaTabla = 1
Else
    VerificaTabla = -1
End If
'Co.CierraConexion
'Set Co = Nothing
Set rs = Nothing
End Function

Public Function GetServerConsol() As String
Dim rs As New ADODB.Recordset

Set rs = coConex.CargaRecordSet("select nConsSisValor from ConstSistema where nConsSisCod=43")
GetServerConsol = IIf(IsNull(rs!nConsSisValor), "", rs!nConsSisValor)
Set rs = Nothing

End Function

Public Function ObtenerPersonas(ByVal psPersCod As String) As ADODB.Recordset

Dim SQL As String

SQL = "SELECT CPERSCOD, CPERSNOMBRE, 0 AS nEndeudaDol  FROM PERSONA  WHERE CPERSCOD ='" & psPersCod & "'"

Set ObtenerPersonas = coConex.CargaRecordSet(SQL)

End Function
'MIOL 20120801, SEGUN RQ12077 ****************************************************
Public Function getCreditosSugeridosHipotecarios(Optional ByVal pnTipoRegCtrl As TipoRegControl = gTpoRegCtrlInformeLegal, Optional psCtaCod As String = "", _
                                                 Optional ByVal psTitular As String = "", Optional ByVal psAgeCod As String = "") As ADODB.Recordset   'OK
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    lsSQL = "exec stp_sel_CredSugHipotecarios '" & psCtaCod & "','" & psTitular & "','" & psAgeCod & "'," & pnTipoRegCtrl
    Set rs = oCom.CargaRecordSet(lsSQL)
    Set getCreditosSugeridosHipotecarios = rs
    oCom.CierraConexion
Set oCom = Nothing
End Function

Public Function getMontoVinculados(ByVal pscCtaCod As String, ByVal pdFecSist As String) As ADODB.Recordset 'OK
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    lsSQL = "exec stp_sel_VerificaSaldosVinculadosRelacionadosxTitular '" & pscCtaCod & "','" & pdFecSist & "'"
    Set rs = oCom.CargaRecordSet(lsSQL)
    Set getMontoVinculados = rs
    oCom.CierraConexion
Set oCom = Nothing
End Function

Public Function getCreditoAprobadosTotalesSC() As ADODB.Recordset 'OK
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    lsSQL = "exec stp_sel_CredAprobadosTotalSC"
    Set rs = oCom.CargaRecordSet(lsSQL)
    Set getCreditoAprobadosTotalesSC = rs
    oCom.CierraConexion
Set oCom = Nothing
End Function

Public Function getParametrosRevision() As ADODB.Recordset ' OK
Dim SQL As String

SQL = "select nParValor from Parametro where nParCod=1004 and nParProd=1000"

Set getParametrosRevision = coConex.CargaRecordSet(SQL)
End Function

Public Function insEstadosExpediente(ByVal pscCtaCod As String, ByVal pscEstUbicacion As String, _
                                                 ByVal pscIngExpediente As String, ByVal pscSalObsExpediente As String, _
                                                 ByVal pscReIngObsExpediente As String, ByVal pscSalExpediente As String, _
                                                 ByVal psnEstRevExpediente As Integer, ByVal psnPrdEstado As Integer, _
                                                 Optional ByVal pnTpoRegControl As TipoRegControl = gTpoRegCtrlRiesgos)  'OK
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    'lsSQL = "exec stp_ins_ColocacRevisionExp '" & pscCtaCod & "','" & pscEstUbicacion & "','" & pscIngExpediente & "','" & pscSalObsExpediente & "','" & pscReIngObsExpediente & "','" & pscSalExpediente & "','" & psnEstRevExpediente & "','" & psnPrdEstado & "'" 'RECO20161018 ERS060-2016
    lsSQL = "exec stp_ins_ColocacRevisionExp '" & pscCtaCod & "','" & pscEstUbicacion & "','" & pscIngExpediente & "','" & pscSalObsExpediente & "','" & pscReIngObsExpediente & "','" & pscSalExpediente & "','" & psnEstRevExpediente & "','" & psnPrdEstado & "'," & pnTpoRegControl 'RECO20161018 ERS060-2016
    Set rs = oCom.CargaRecordSet(lsSQL)
    Set insEstadosExpediente = rs
    oCom.CierraConexion
Set oCom = Nothing
End Function

Public Function updateEstadoExpediente(ByVal pscCtaCod As String, Optional ByVal pnTpoRegControl As TipoRegControl = gTpoRegCtrlRiesgos) 'OK BY ARLO MODIFY 20171027
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    lsSQL = "exec stp_update_EstColocacRevisionExp'" & pscCtaCod & "'," & pnTpoRegControl 'BY ARLO MODIFY 20171027
    Set rs = oCom.CargaRecordSet(lsSQL)
    Set updateEstadoExpediente = rs
    oCom.CierraConexion
Set oCom = Nothing
End Function

Public Function updateParametroRevision(ByVal psnParRevExp As Integer) 'OK
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    lsSQL = "exec stp_update_ParametroRevisionCred'" & psnParRevExp & "'"
    Set rs = oCom.CargaRecordSet(lsSQL)
    Set updateParametroRevision = rs
    oCom.CierraConexion
Set oCom = Nothing
End Function
'END MIOL ***********************************************************************

'MIOL 20120920, SEGUN RQ12257 ***************************************************
'OK
Public Function getSaldoClienteSunat(ByVal pscNroDoc As String, ByVal psnTipo As Integer) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    If psnTipo = 1 Then
        lsSQL = "exec stp_sel_SaldoClienteSunatAhorros '" & pscNroDoc & "'"
    ElseIf psnTipo = 2 Then
        lsSQL = "exec stp_sel_SaldoClienteSunatPlazoFijo '" & pscNroDoc & "'"
    End If
    Set rs = oCom.CargaRecordSet(lsSQL)
    Set getSaldoClienteSunat = rs
    oCom.CierraConexion
Set oCom = Nothing
End Function

'OK
Public Function getGaranSaldoClienteSunat(ByVal pscNroDoc As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
        lsSQL = "exec stp_sel_GaranSaldoClientexNroDoc '" & pscNroDoc & "'"
    Set rs = oCom.CargaRecordSet(lsSQL)
    Set getGaranSaldoClienteSunat = rs
    oCom.CierraConexion
Set oCom = Nothing
End Function
'END MIOL ***********************************************************************

'WIOR 20160623 ***
Public Function SobreEndVerificaCartera() As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_SobreEndCodCarteraFecha"
    oCon.AbreConexion
    Set SobreEndVerificaCartera = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Sub SobreEndEjecutarCartera()
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_ins_SobreEndCodCarteraEjecutar"
    oCon.AbreConexion
    oCon.Ejecutar sSql
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'WIOR FIN ********
'RECO20161020 ERS060-2016***********************************************
Public Function ValidaExisteRegProceso(ByVal psCtaCod As String, ByVal pnTipoRegControl As TipoRegControl) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ValidaExisteRegProceso '" & psCtaCod & "'," & pnTipoRegControl
    oCon.AbreConexion
    Set ValidaExisteRegProceso = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ListaCreditosSeguimiento(ByVal psCtaCod As String, ByVal psPersNombre As String, ByVal psAgeCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_ListaCreditosSeguimiento '" & psCtaCod & "','" & psPersNombre & "','" & psAgeCod & "'"
    oCon.AbreConexion
    Set ListaCreditosSeguimiento = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function CredDetalleSeguimiento(ByVal psCtaCod As String) As ADODB.Recordset
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_sel_CredDetalleSeguimiento '" & psCtaCod & "'"
    oCon.AbreConexion
    Set CredDetalleSeguimiento = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
    Set oCon = Nothing
End Function
'RECO FIN***************************************************************
'ARLO 20170919, SEGUN ERS060-2016 ****************************************************
Public Function getCreditosAprobadosMinuta(Optional ByVal pnTipoRegCtrl As TipoRegControl = gTpoRegCtrlInformeLegal, Optional psCtaCod As String = "", _
                                                 Optional ByVal psTitular As String = "", Optional ByVal psAgeCod As String = "") As ADODB.Recordset   'OK
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    lsSQL = "exec stp_sel_CredAprobadosMinuta '" & psCtaCod & "','" & psTitular & "','" & psAgeCod & "'," & pnTipoRegCtrl
    Set rs = oCom.CargaRecordSet(lsSQL)
    Set getCreditosAprobadosMinuta = rs
    oCom.CierraConexion
Set oCom = Nothing
End Function
'ARLO 20170919, SEGUN ERS060-2016 ****************************************************
Public Function getCreditosAsistenteAgencia(Optional ByVal pnTipoRegCtrl As TipoRegControl = gTpoRegCtrlInformeLegal, Optional psCtaCod As String = "", _
                                                 Optional ByVal psTitular As String = "", Optional ByVal psAgeCod As String = "") As ADODB.Recordset   'OK
Dim rs As New ADODB.Recordset
Dim oCom As COMConecta.DCOMConecta
Dim lsSQL As String
Set oCom = New COMConecta.DCOMConecta
    oCom.AbreConexion
    lsSQL = "exec stp_sel_CreditosAsistenteAgencia '" & psCtaCod & "','" & psTitular & "','" & psAgeCod & "'," & pnTipoRegCtrl
    Set rs = oCom.CargaRecordSet(lsSQL)
    Set getCreditosAsistenteAgencia = rs
    oCom.CierraConexion
Set oCom = Nothing
End Function
'CTI1 20180817 ***
Public Sub CalculoCapPagoSobreEndEjecutarCartera(ByVal pdFecha As Date)
    Dim sSql As String
    Dim oCon As New COMConecta.DCOMConecta
    sSql = "EXEC stp_ins_GenerarCapacidadPagoACierre '" & Format(pdFecha, "YYYYMMDD") & "'"
    
    oCon.AbreConexion
    oCon.Ejecutar sSql
    oCon.CierraConexion
    Set oCon = Nothing
End Sub
'CTI1 FIN ********
