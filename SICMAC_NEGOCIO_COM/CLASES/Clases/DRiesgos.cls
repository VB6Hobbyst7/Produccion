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
Dim Sql As String
Dim Co As DConecta
Set Co = New DConecta

Sql = "select NCONSVALOR, CCONSDESCRIPCION"
Sql = Sql & " from constante WHERE NCONSCOD = 9060 AND NCONSVALOR != 9060"
Sql = Sql & " ORDER BY NCONSVALOR"

Co.AbreConexion
Set TiposGaratiaRiesgos = Co.CargaRecordSet(Sql)
Co.CierraConexion
Set Co = Nothing
End Function

Function DiaCierreCalifMes(ByVal psServer As String) As String
Dim Sql As String
Dim rs As ADODB.Recordset
Dim Co As DConecta

Set Co = New DConecta

Co.AbreConexion
Sql = "select nConsSisValor from constsistema where nConsSisCod = 160"
Set rs = Co.CargaRecordSet(Sql)
Co.CierraConexion
DiaCierreCalifMes = "-1"
If Not (rs.EOF And rs.BOF) Then
    DiaCierreCalifMes = Format(rs!NCONSSISVALOR, "DD/MM/YYYY")
Else
    MsgBox "No exite el codigo 155, avise a Sistemas", vbInformation, "AVISO"
End If
Set rs = Nothing
Set Co = Nothing
End Function
 
Function FechaCalificacion(ByVal psServer As String, ByVal Fecha As String, ByVal PersCod As String) As ADODB.Recordset
Dim Sql As String
Dim rs As ADODB.Recordset
Dim Co As DConecta

Sql = "select dfecha, cCtaCod, nDiasAtraso,"
Sql = Sql & " isnull(cCalGen,'-') Calgen,  isnull(cCalSistF,'-') CalSis,"
Sql = Sql & " isnull(cCalCMAC,'-') CMAC, isnull(cCalEval,'-') CalHist"
Sql = Sql & " From " & psServer & "ColocCalifProvTotal"
Sql = Sql & " where datediff(month,dfecha,'" & Fecha & "')=0"
Sql = Sql & " and cPersCod =  '" & PersCod & "'"

Set Co = New DConecta
Co.AbreConexion
Set FechaCalificacion = Co.CargaRecordSet(Sql)
Co.CierraConexion

Set Co = Nothing
Set rs = Nothing

End Function

Function GetCreditodMesReclasificados(ByVal psServer As String, ByVal nTipoCambio As Double) As ADODB.Recordset
Dim Sql As String
Dim Co As DConecta

Sql = " Select cPersCod, cPersNombre,"
Sql = Sql & " Isnull(Sum(case when substring(Rd.cod_cuenta,3,1) = '2' And RD.cod_cuenta like '14_[1456]%' then Val_saldo / " & nTipoCambio & " End),0) nEndeudaDol"
Sql = Sql & " From"
Sql = Sql & " ("
Sql = Sql & "     Select cPersNombre,CP.cPersCod, case when DNI =  null Or DNI ='' then Ruc else DNI end Doc from"
Sql = Sql & "     (Select distinct C.cPersCod From ColocCalifProv C where nPrdEstado in (2020,2021,2022,2030,2031,2032) and substring(cCtaCod,6,1)='2') CP"
Sql = Sql & "     Inner Join"
Sql = Sql & "     (SELECT cPersCod , cPersNombre FROM PERSONA) P on P.cPersCod = Cp.cPersCod"
Sql = Sql & "     Left Join"
Sql = Sql & "      (Select cPersCod,cPersIDnro DNI from PersID where cPersIdTpo='1') PDNI on PDNI.cPersCod = CP.cPersCod"
Sql = Sql & "     Left Join"
Sql = Sql & "      (Select cPersCod,cPersIDnro RUC from PersID where cPersIdTpo='2') PRUC on PRuc.cPersCod = CP.cPersCod"
Sql = Sql & " ) Per"
Sql = Sql & " Inner Join " & psServer & "RccTotal R on LTRIM(RTRIM(R.Cod_Doc_Id)) = LTRIM(RTRIM(Per.Doc))"
Sql = Sql & " Inner Join " & psServer & "RccTotalDet RD on LTRIM(RTRIM(R.Cod_Edu)) = LTRIM(RTRIM(RD.Cod_Edu))"
Sql = Sql & " WHERE ( LTRIM(RTRIM(Per.doc)) <>'' or Per.doc <>null ) And"
Sql = Sql & " RD.cod_cuenta like '14_[1456]%' and nEndeudaDol > 30000 "
Sql = Sql & " Group by cPersCod,cPersNombre"

Set Co = New DConecta
Co.AbreConexion
Set GetCreditodMesReclasificados = Co.CargaRecordSet(Sql)
Co.CierraConexion
Set Co = Nothing
End Function

Function NuevoCredReclasificado(ByVal psFecSis As String, ByVal PersCod As String, ByVal psCodCmac As String, _
ByVal psgColComercEmp As Integer, ByVal psMovNro As String, _
ByVal psCodAge As String, ByVal psCodUser As String) As Integer
Dim Sql As String
Dim Co As DConecta
Dim Mo As DMov
Dim rs As ADODB.Recordset
Dim psNuevaCta As String
Dim psCuenta As String
Dim psMoneda As String
Dim oGen As DGeneral

Dim NroCred As Integer
Dim cMovNro As String

Dim Sql0, Sql1, Sql2, Sql3, Sql4, Sql5, Sql6, Sql7, Sql8, Sql9, Sql10, Sql11, Sql12, Sql13, Sql14 As String
Dim Sql15, Sql16 As String
Set Co = New DConecta
Set oGen = New DGeneral
Set Mo = New DMov

Dim nMovNro As Long
Sql = "Select cCtaCod, cPersCod from ProductoPersona where substring(cCtaCod,6,1)='2'"
Sql = " and nPrdPersRelac = 20 and cPersCod = '" & PersCod & "'"

On Error GoTo ErrorReclasificaMesComercial

Co.AbreConexion
Co.BeginTrans
Set rs = Co.CargaRecordSet(Sql)
NroCred = 0
While Not rs.EOF
    NroCred = NroCred + 1
    cMovNro = Mo.GeneraMovNro(gdFecSis, psCodAge, psCodUser)
    nMovNro = Mo.InsertaMov(psMovNro, 172101, "Reclasificacion Mes Comercial")
    psCuenta = rs!cCtaCod
    psMoneda = Mid(rs!cCtaCod, 9, 1)
    psNuevaCta = psCodCmac & oGen.GeneraNuevaCuenta(Mid(rs!cCtaCod, 4, 2), psgColComercEmp, Mid(rs!cCtaCod, 6, 1))
        
    Sql15 = "INSERT INTO Producto(cCtaCod,nTasaInteres, nSaldo,nPrdEstado,dPrdEstado,nTransacc)"
    Sql15 = Sql15 & " Select  '" & psNuevaCta & "' cCtaCod,nTasaInteres, nSaldo,nPrdEstado,dPrdEstado,nTransacc"
    Sql15 = Sql15 & "From producto where cCtaCod = '" & psCuenta & "'"
    Co.Ejecutar (Sql15)
    
    Sql14 = "Insert ColocReclasificaMes"
    Sql14 = Sql14 & " (dFecha,cCtaCod,cCtaCodAnt,nSaldoCap,cCodope,nMovActualiza) values"
    Sql14 = Sql14 & "('" & Format(psFecSis, "YYYY/MM/DD") & "','" & psNuevaCta & "','" & psCuenta & "',"
    Sql14 = Sql14 & "'172101'," & nMovNro & "')"
    Co.Ejecutar (Sql14)
    
    Sql0 = " Update ProductoPersona set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Sql1 = " Update Colocaciones set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Sql2 = " Update colocacCred set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Sql3 = " Update ColocacEstado set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Sql4 = " Update ColocCalendario set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Sql5 = " Update ColocCalendDet set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Sql6 = " Update ColocCalificacionAnalista set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Sql7 = " Update ColocCredCredVig set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Sql8 = " Update ColocGarantia set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Sql9 = " Update ColocGarantRec set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Sql10 = " Update movCol set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    Sql11 = " Update MovColDet set cCtaCod='" & psNuevaCta & "'  where cCtaCod='" & psCuenta & "'"
    
    Sql16 = "Delete Producto where cCtaCod = '" & psCuenta & "'"
    
    Co.Ejecutar (Sql0):    Co.Ejecutar (Sql1)
    Co.Ejecutar (Sql2):    Co.Ejecutar (Sql3)
    Co.Ejecutar (Sql4):    Co.Ejecutar (Sql5)
    Co.Ejecutar (Sql6):    Co.Ejecutar (Sql7)
    Co.Ejecutar (Sql8):    Co.Ejecutar (Sql9)
    Co.Ejecutar (Sql10):    Co.Ejecutar (Sql11)
    Co.Ejecutar (Sql16)
    
    rs.MoveNext
Wend
Co.CommitTrans
Co.CierraConexion
'Codigo de reclasificacion Mes Comercial 172101
    Set Co = Nothing
    Set oGen = Nothing
    Set Mo = Nothing
    Set rs = Nothing
    NuevoCredReclasificado = NroCred
    
Exit Function
ErrorReclasificaMesComercial:
    Co.RollbackTrans
    Set Co = Nothing
    Set oGen = Nothing
    Set Mo = Nothing
    Set rs = Nothing
    NuevoCredReclasificado = -1
Call RaiseError(MyUnhandledError, "DRiesgo:NuevoCredReclasificado Method")
End Function

