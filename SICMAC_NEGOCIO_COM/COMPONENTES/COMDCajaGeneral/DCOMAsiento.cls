VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMAsiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim gnTipCambio As Currency
Dim gnTipCambioV As Currency
Dim gnTipCambioC As Currency


'************ TRABAJOS CON EL ASIENTO AVMM  ***********************************
'*********************************************************************
Public Function GetAsientoParametro(pnAsiParCod As Long) As String
    Dim sql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    
    sql = "Select IsNull(cAsiParValor,'') Valor From AsientoParametro Where nAsiParCod = " & pnAsiParCod
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        GetAsientoParametro = ""
    Else
        GetAsientoParametro = rs!Valor
    End If
    oCon.CierraConexion
End Function

Public Sub InsertaAsientoValida(ByVal pdFecha As Date, ByVal pscoduser As String, ByVal pdHoraGrab As Date)
    Dim ssql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim oFunV As New COMFunciones.FCOMVarPublicas
    
    oCon.AbreConexion
    
    ssql = " INSERT INTO AsientoValida (dAsientoFecha, cAsientoTipo, cAsientoEstado, cCodUsu, dAsientoModif)" & _
           " VALUES ('" & Format(pdFecha, oFunV.gsFormatoFecha) & "','2','0','" & pscoduser & "','" & Format(pdHoraGrab, oFunV.gsFormatoFechaHora) & "') "
    oCon.Ejecutar ssql
    
    oCon.CierraConexion

End Sub

Public Sub EliminarAsiento(ByVal pdFecha As Date)
    Dim ssql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim oFunV As New COMFunciones.FCOMVarPublicas
    
    oCon.AbreConexion
    
    ssql = "EXEC EliminaAsientoDB '" & Format(pdFecha, oFunV.gsFormatoFecha) & "'"
    oCon.Ejecutar ssql
    
    oCon.CierraConexion

End Sub

Public Function ObtenerNroAsientos(ByVal pdFecha As Date) As ADODB.Recordset
    Dim ssql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim oFunV As New COMFunciones.FCOMVarPublicas
    
    oCon.AbreConexion
    
    ssql = "SELECT ISNULL(COUNT(*),0) as nRegAsiento FROM ASIENTODN WHERE CONVERT(CHAR(10),dFecha,112)='" & Format(pdFecha, "yyyymmdd") & "' "
    Set rs = oCon.CargaRecordSet(ssql)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerNroAsientos = rs
    Else
        Set ObtenerNroAsientos = Nothing
    End If
    oCon.CierraConexion

End Function

Public Sub InsertaAsientoDN(ByVal pdHoraGrab As Date, ByVal psvCodConta As String, ByVal pnMontoD As Double, ByVal pnMontoH As Double, _
                            Optional ByVal pstipo As String, Optional ByVal psvagencia As String, Optional ByVal pscOpeCod As String, Optional ByVal pscCtaCod As String, Optional psnMov As String)
    Dim sql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim oFunV As New COMFunciones.FCOMVarPublicas
    
    oCon.AbreConexion
    If Trim(psnMov) <> "" Then
        If Trim(pscCtaCod) <> "" Then
            sql = " INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,nMovNro,cOpeCod, cCtaCod ) " & _
                  " VALUES('" & Format(pdHoraGrab, "yyyymmdd hh:mm:ss") & "','" & psvCodConta & "'," & pnMontoD & "," & pnMontoH & ",'" & pstipo & "','" & psvagencia & "'," & psnMov & ",'" & pscOpeCod & "','" & pscCtaCod & "')"
        Else
            sql = " INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,nMovNro,cOpeCod ) " & _
                  " VALUES('" & Format(pdHoraGrab, "yyyymmdd hh:mm:ss") & "','" & psvCodConta & "'," & pnMontoD & "," & pnMontoH & ",'" & pstipo & "','" & psvagencia & "'," & psnMov & ",'" & pscOpeCod & "')"
        End If
    Else
        If Trim(pscOpeCod) <> "" Then
            sql = " INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge,cOpeCod, cCtaCod ) " & _
                  " VALUES('" & Format(pdHoraGrab, "yyyymmdd hh:mm:ss") & "','" & psvCodConta & "'," & pnMontoD & "," & pnMontoH & ",'" & pstipo & "','" & psvagencia & "','" & pscOpeCod & "','" & pscCtaCod & "')"
        Else
            sql = " INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
                  " VALUES('" & Format(pdHoraGrab, "yyyymmdd hh:mm:ss") & "','" & psvCodConta & "'," & pnMontoD & "," & pnMontoH & ",'" & pstipo & "','" & psvagencia & "')"
        End If
    End If
    oCon.Ejecutar sql
    
    oCon.CierraConexion

End Sub

Public Function GetPlantillaPuente(ByVal psPlantillaAnt As String, ByVal lsProducto As String, ByVal pnConcepto As Long, ByVal psOpeCod As String) As String
Dim oCon As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset
Dim sql As String
Set oCon = New COMConecta.DCOMConecta

GetPlantillaPuente = ""
oCon.AbreConexion
sql = "SELECT   cCtaContCodNew FROM PlantillaAsientoPuente " _
    & " where   cCtaContCodAnt = '" & psPlantillaAnt & "' and cProducto ='" & lsProducto & "' " _
    & "         and nPrdConceptoCod = " & pnConcepto & " AND cOpeCod ='" & psOpeCod & "'"
    
Set rs = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing
If Not rs.EOF And Not rs.BOF Then
    GetPlantillaPuente = Trim(rs!cCtaContCodNew)
End If
rs.Close
Set rs = Nothing

End Function

Public Function ObtenerSumaMontoN(ByVal pdHoraGrab As Date) As ADODB.Recordset
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    sql = " SELECT round(SUM(ndebe),2) AS MonDebe , round(SUM(nhaber),2) AS MonHaber FROM AsientoDN " & _
          " WHERE ctipo = '0' AND substring(cCtaCnt,3,1) = '" & COMDConstantes.gMonedaNacional & "' AND datediff(dd, dFecha ,'" & Format(pdHoraGrab, "yyyymmdd hh:mm:ss") & "') = 0 " & _
          " Group BY cCtaCnt "
            
    Set rs = oCon.CargaRecordSet(sql)
    If (rs.BOF Or rs.EOF) Then
        Set ObtenerSumaMontoN = Nothing
    Else
        Set ObtenerSumaMontoN = rs
    End If
    
    oCon.CierraConexion
    
End Function

Public Function VarAsientoProdEquiv(ByVal pProducto As String, ByVal pEquivTipo As String) As String
    Dim RegCta As New ADODB.Recordset
    Dim tmpSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    tmpSql = "SELECT cEquivAsiento AS Campo FROM AsientoProdEquiv " & _
                "WHERE cProducto ='" & pProducto & "' AND cEquivTpo = '" & pEquivTipo & "' "
    
    Set RegCta = oCon.CargaRecordSet(tmpSql)
    If (RegCta.BOF Or RegCta.EOF) Then
        VarAsientoProdEquiv = ""
    Else
        VarAsientoProdEquiv = IIf(IsNull(RegCta!Campo), "", Trim(RegCta!Campo))
    End If
    RegCta.Close
    Set RegCta = Nothing
End Function

Function GetAgenciaCtaDesmAbonoCta(ByVal lnMovNro As Long, ByVal lsOpeCod As String) As String
    Dim oCon As COMConecta.DCOMConecta
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    Set oCon = New COMConecta.DCOMConecta
    GetAgenciaCtaDesmAbonoCta = ""
    oCon.AbreConexion
    sql = "select SUBSTRING(cCtaCod,4,2) as cCodAge from movcap where nmovnro =" & lnMovNro & " and copecod ='" & lsOpeCod & "'"
    Set rs = oCon.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        GetAgenciaCtaDesmAbonoCta = rs!cCodAge
    End If
    oCon.CierraConexion
End Function

Public Function AsientoParche(pAsiento As String, Optional ByVal pNuePlan As Boolean = False) As String
    Dim tmpRs As New ADODB.Recordset
    Dim tmpSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    
    If pNuePlan Then
        tmpSql = " SELECT cCodCntNue" & _
            " FROM AsientoPuenteN " & _
            " WHERE cCodCnt = '" & Trim(pAsiento) & "'"
    Else
        tmpSql = " SELECT cCodCntNue" & _
            " FROM AsientoPuente " & _
            " WHERE cCodCnt = '" & Trim(pAsiento) & "'"
    End If
    Set tmpRs = oCon.CargaRecordSet(tmpSql)
    If (tmpRs.BOF Or tmpRs.EOF) Then
        AsientoParche = ""
    Else
        AsientoParche = Trim(tmpRs!cCodCntNue)
    End If
    Set oCon = Nothing
End Function

Public Function CuadreCtas(ByVal pdHoraGrab As Date, ByVal psNroCta As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim tmpSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    
    tmpSql = "SELECT abs(round(SUM(ndebe),2)) nDebe,  abs(round(SUM(nhaber),2)) nHaber FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & Format(pdHoraGrab, "yyyymmdd hh:mm:ss") & "') = 0 " & _
            " AND cCtaCnt like '" & psNroCta & "'" & _
            " Group BY cCtaCnt "
        
    Set rs = oCon.CargaRecordSet(tmpSql)
    If (rs.BOF Or rs.EOF) Then
        Set CuadreCtas = Nothing
    Else
        Set CuadreCtas = rs
    End If
    Set oCon = Nothing
    
End Function

Public Function AsientosCompraVentaME(ByVal pdFechaAsiento As String, ByVal gAsientoProcesoSiAsientoCVME As String) As ADODB.Recordset
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim oFunV As New COMFunciones.FCOMVarPublicas
    oCon.AbreConexion
      sql = " SELECT M.nMovNro, M.cMovNro, cOpeCod, nMovImporte AS Monto, MTC.nMovTpoCambio nTipCambio, " _
           & " SUBSTRING(M.cMovNro,18,2) Agencia, round(nMovImporte*nMovTpoCambio,2) AS NMOVSOLES  FROM Mov M" _
           & " Inner Join MovCompraVenta MCV On M.nMovNro = MCV.nMovNro" _
           & " Inner Join MovTpoCambio MTC On M.nMovNro = MTC.nMovNro" _
           & " WHERE cOpeCod IN (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoSiAsientoCVME & ")  AND (M.nMovFlag = " & COMDConstantes.gMovFlagVigente & ")" _
           & " And M.cMovNro Like '" & Format(pdFechaAsiento, oFunV.gsFormatoMovFecha) & "%' AND MCV.nMovImporte <> 0 ORDER BY cOpeCod"
           
    Set rs = oCon.CargaRecordSet(sql)
    If (rs.BOF Or rs.EOF) Then
        Set AsientosCompraVentaME = Nothing
    Else
        Set AsientosCompraVentaME = rs
    End If
    Set oCon = Nothing
End Function

Public Function ValidarDolares(ByVal pdHoraGrab As Date) As ADODB.Recordset
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    sql = " SELECT cCodAge, sum(round(ndebe,2)) nDebe , sum(round(nhaber,2)) nHaber " & _
            " FROM AsientodN " & _
            " WHERE cTipo = '1' AND datediff(dd, dFecha ,'" & Format(pdHoraGrab, "yyyymmdd hh:mm:ss") & "') = 0 " & _
            " GROUP BY cCodAge, cCtaCnt Order by 1 "
    
    Set rs = oCon.CargaRecordSet(sql)
    If (rs.BOF Or rs.EOF) Then
        Set ValidarDolares = Nothing
    Else
        Set ValidarDolares = rs
    End If
    oCon.CierraConexion
    
End Function

Public Function ObtenerSumaDieferenciaE(ByVal pdHoraGrab As Date) As ADODB.Recordset
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    
    sql = "SELECT round(SUM(ndebe),2) AS MonDebe , round(SUM(nhaber),2) AS MonHaber FROM AsientoDN " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & Format(pdHoraGrab, "yyyymmdd hh:mm:ss") & "') = 0 " & _
            " Group BY cCtaCnt "
    
    Set rs = oCon.CargaRecordSet(sql)
    If (rs.BOF Or rs.EOF) Then
        Set ObtenerSumaDieferenciaE = Nothing
    Else
        Set ObtenerSumaDieferenciaE = rs
    End If
    
    oCon.CierraConexion
    
End Function

Public Sub ActualizaAsientoValida(ByVal pAsiDia As Boolean, ByVal pscoduser As String, ByVal pdFecha As String, ByVal pdHoraGrab As Date)
    Dim sql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim oFunV As New COMFunciones.FCOMVarPublicas
    
    oCon.AbreConexion
    
    If pAsiDia Then
        'Actualización de variable de control
        sql = "UPDATE ConstSistema SET nConsSisValor = '1' WHERE nConsSisCod = '2'"
        oCon.Ejecutar sql
    End If
    
    sql = "UPDATE AsientoValida SET cAsientoEstado = '1' " & _
         " WHERE dAsientoFecha = '" & Format(pdFecha, oFunV.gsFormatoFecha) & "' AND cAsientoTipo = '2' AND " & _
         " cCodUsu = '" & pscoduser & "' AND dAsientoModif = '" & Format(pdHoraGrab, oFunV.gsFormatoFechaHora) & "'"
    oCon.Ejecutar sql
    
    oCon.CierraConexion
    
End Sub

Public Function ObtenerSumaMNE(ByVal pstipo As String, ByVal pnMN As Integer, ByVal pnME As Integer, ByVal pdHoraGrab As Date, Optional ByVal vcad As String) As ADODB.Recordset
    Dim ssql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    
    ssql = " SELECT  cctacnt, round(sum(ndebe),2) nDebe, round(sum(nhaber),2) nHaber, substring(cctacnt,3,1) AS MONEDA " & _
            " FROM AsientoDN "
    If Trim(pstipo) = "3" Then
        ssql = ssql & " WHERE cTipo IN ('" & pstipo & "') AND datediff(dd,dFecha,'" & Format(pdHoraGrab, "yyyymmdd hh:mm:ss") & "') = 0  And substring(cctacnt,3,1) = '2'" & _
                           " GROUP BY CTIPO, CCTACNT ORDER BY cctacnt "
    Else
        ssql = ssql & " WHERE cTipo IN ('" & pstipo & "') AND substring(cctacnt,3,1) IN (" & vcad & ") AND datediff(dd,dFecha,'" & Format(pdHoraGrab, "yyyymmdd hh:mm:ss") & "') = 0  "
    End If
    If Trim(pstipo) <> "3" Then
        If pnMN = 1 And pnME = 1 Then
            ssql = ssql & " GROUP BY CTIPO, CCTACNT ORDER BY moneda , cctacnt "
        Else
            ssql = ssql & " GROUP BY CTIPO, CCTACNT ORDER BY cctacnt "
        End If
    End If
    Set rs = oCon.CargaRecordSet(ssql)
    If (rs.BOF Or rs.EOF) Then
        Set ObtenerSumaMNE = Nothing
    Else
        Set ObtenerSumaMNE = rs
    End If
    
    oCon.CierraConexion
End Function

Public Function CompraVentaMEN(ByVal pdHoraGrab As Date) As ADODB.Recordset
     Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    
    sql = " SELECT  cctacnt, round(sum(ndebe),2) nDebe, round(sum(nhaber),2) nHaber, ctipo " & _
           " FROM AsientoDN WHERE ctipo IN ('1','2') AND datediff(dd, dFecha ,'" & Format(pdHoraGrab, "yyyymmdd hh:mm:ss") & "') = 0 " & _
           " GROUP BY CTIPO, CCTACNT "
        
    Set rs = oCon.CargaRecordSet(sql)
    If (rs.BOF Or rs.EOF) Then
        Set CompraVentaMEN = Nothing
    Else
        Set CompraVentaMEN = rs
    End If
    
    oCon.CierraConexion
End Function

Public Function CtaAsiento(ByVal pFecha As Date, ByVal pTipo As String, ByVal pMoneda As String, _
Optional ByVal pOtroTipo As String = "", Optional ByVal pNuePlan As Boolean = False, Optional psAgeCod As String = "") As Currency
    Dim tmpReg As ADODB.Recordset
    Dim tmpSql As String, ssql As String
    Set tmpReg = New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    
    tmpReg.CursorLocation = adUseClient
    If pNuePlan Then
        tmpSql = "SELECT (sum(ndebe) - sum(nhaber)) Campo FROM AsientoDN " & _
            " WHERE datediff(dd,dfecha,'" & Format(pFecha, "mm/dd/yyyy") & "') = 0 "
        If pTipo = "A" Then         ' Caja
'            If gbAgeEsp Then
'                sSql = " AND substring(cCtaCnt,1,6) IN ('11" & pMoneda & "103')"
'            Else
'                sSql = " AND substring(cCtaCnt,1,6) IN ('11" & pMoneda & "102')"
'            End If
            ssql = " AND Substring(cCtaCnt,1,6) IN ('11" & pMoneda & "102','11" & pMoneda & "103')"
            
            tmpSql = "SELECT SUM(nDebe - nHaber) Campo From ( " _
                & "SELECT cTipo, Round(SUM(nDebe),2) nDebe, Round(SUM(nHaber),2) nHaber FROM AsientoDN " _
                & "WHERE datediff(dd,dfecha,'" & Format(pFecha, "mm/dd/yyyy") & "') = 0 " & ssql & " " _
                & IIf(psAgeCod <> "", " and cCodAge = '" & psAgeCod & "' ", "") & "GROUP BY cTipo) T"
        ElseIf pTipo = "B" Then     ' Créditos
            'tmpSql = tmpSql & " AND substring(cCtaCnt,1,3) IN ('14" & pMoneda & "') " & _
            '    " AND (substring(cCtaCnt,1,4) Not IN ('1419','1429','1416','1426') " & _
            '    " AND (substring(cCtaCnt,1,4) Not IN ('14" & pMoneda & "5') OR " & _
            '    " substring(cCtaCnt,9,1) Not IN ('9','2'))) "
                
            tmpSql = tmpSql & " AND substring(cCtaCnt,1,3) IN ('14" & pMoneda & "') "
            tmpSql = tmpSql & "     AND ( substring(cCtaCnt,1,4) Not IN ('1419','1429','1418','1428') "
            tmpSql = tmpSql & " ) "
            tmpSql = tmpSql & IIf(psAgeCod <> "", " and SubString(cCtaCod,4,2) = '" & psAgeCod & "' ", "")
        ElseIf pTipo = "C" Then     ' Ahorros
            'Se agrega la 2117 y 2127 por las inmovilizadas forman parte del saldo
            tmpSql = tmpSql & " AND substring(cCtaCnt,1,4) IN ('21" & pMoneda & "2','21" & pMoneda & "3','23" & pMoneda & "2','23" & pMoneda & "3','21" & pMoneda & "7')"
            tmpSql = tmpSql & IIf(psAgeCod <> "", " and SubString(cCtaCod,4,2) = '" & psAgeCod & "' ", "")
       
        End If
    Else
        tmpSql = "SELECT (sum(ndebe) - sum(nhaber)) Campo FROM AsientoD " & _
            " WHERE datediff(dd,dfecha,'" & Format(pFecha, "mm/dd/yyyy") & "') = 0 "
        If pTipo = "A" Then
            tmpSql = tmpSql & " AND substring(cCtaCnt,1,6) IN ('11" & pMoneda & "103')"
        ElseIf pTipo = "B" Then
            tmpSql = tmpSql & " AND substring(cCtaCnt,1,3) IN ('14" & pMoneda & "') " & _
                " AND substring(cCtaCnt,1,4) Not IN ('1419','1429','1416','1426')"
        ElseIf pTipo = "C" Then
            tmpSql = tmpSql & " AND substring(cCtaCnt,1,3) IN ('23" & pMoneda & "','24" & pMoneda & "','26" & pMoneda & "')"
        End If
    End If
    tmpSql = tmpSql & IIf(Len(Trim(pOtroTipo)) = 0, " And cTipo IN ('0')", " Where cTipo IN ('0','" & pOtroTipo & "')")
    Set tmpReg = oCon.CargaRecordSet(tmpSql)
    Set tmpReg.ActiveConnection = Nothing
    If (tmpReg.BOF Or tmpReg.EOF) Then
        CtaAsiento = 0
    Else
        CtaAsiento = IIf(IsNull(tmpReg!Campo), 0, tmpReg!Campo)
    End If
    tmpReg.Close
    Set tmpReg = Nothing
    Set oCon = Nothing
    
End Function

'*********************** CONSULTA DE MOVIMIENTOS- CAPTACIONES -CREDITOS - OPERACIONES *******************************
'********************************************************************************************************************
Public Function CargaTranDiaria(ByVal gAsientoProcesoNoAsientoCab3 As String, ByVal gAsientoProcesoNoAsientoCab4 As String, ByVal gAsientoProcesoNoAsiento As String, ByVal gAsientoProcesoCreditoSinAsiento As String, _
                                ByVal gAsientoProcesoSiAsiento As String, ByVal pdFechaAsiento As Date, ByVal gAsientoProcesoSiColocCargoCuenta As String) As ADODB.Recordset
   Dim sql As String
   Dim rs As New ADODB.Recordset
   Dim oCon As COMConecta.DCOMConecta
   Set oCon = New COMConecta.DCOMConecta
   Dim oFunV As New COMFunciones.FCOMVarPublicas
  
   oCon.AbreConexion

   sql = " SELECT    MC.nCredEstado,  M.nMovNro, M.cMovNro, dbo.FechaHoraMov(M.cMovNro) dFecTran, MC.cCtaCod, " _
          & "           MCD.cOpeCod, MCD.nPrdConceptoCod Concepto, round(abs(MCD.nMonto),2) As SumaMonto,  " _
          & "           ISNULL(MCMAC.cPersCod,SubString(m.cMovNro,18,2)) cCodAge, MD.cDocNro, MCD.nNroCuota, Null  nSaldCnt,  " _
          & "           c.cLineaCred As LineaC, CR.cRFA " _
          & "   FROM Mov M " _
          & "   INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro" _
          & "   INNER JOIN MovColDet MCD ON MC.nMovNro = MCD.nMovNro And MC.cOpeCod  = MCD.cOpeCod And MC.cCtaCod  = MCD.cCtaCod and not MCD.nPrdConceptocod IN (2251,2253, 2254,2311,2312,2217) " _
          & "   LEFT  JOIN MovCMAC MCMAC ON M.nMovNro = MCMAC.nMovNro" _
          & "   LEFT  JOIN Colocaciones C ON MC.cCtaCod = C.cCtaCod" _
          & "   LEFT  JOIN colocacCred CR ON MC.cCtaCod = CR.cCtaCod" _
          & "   LEFT  JOIN ColocRecup CJ ON MC.cCtaCod = CJ.cCtaCod" _
          & "   LEFT  JOIN MovDoc MD ON M.nMovNro = MD.nMovNro AND MD.nDocTpo<>23 " _
          & "   Where (M.nMovFlag = 0) And MCD.nMonto <> 0 And M.cMovNro Like '" & Format(pdFechaAsiento, oFunV.gsFormatoMovFecha) & "%'" _
          & "       AND substring(MC.cOpeCod,1,3) Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoNoAsientoCab3 & ")" _
          & "       And substring(M.cOpeCod,1,4) Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoNoAsientoCab4 & ")" _
          & "       And M.cOpeCod Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoNoAsiento & ")" _
          & "       And MC.cOpeCod Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoCreditoSinAsiento & ")" _
          & "       AND substring(M.cOpeCod,1,2) IN (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoSiAsiento & ")" _
          & "   Union"
          
   sql = sql & "   SELECT 0 as nCredEstado,  M.nMovNro, M.cMovNro, dbo.FechaHoraMov(M.cMovNro) dFecTran, MC.cCtaCod, MCD.cOpeCod, MCD.nConceptoCod Concepto, round(abs(MCD.nMonto),2) As SumaMonto,  ISNULL(MCMAC.cPersCod,SubString(m.cMovNro,18,2)) cCodAge, MD.cDocNro, 0 nNroCuota, MC.nSaldoContable  nSaldCnt, Null As LineaC, NULL AS CRFA FROM Mov M" _
          & "   INNER JOIN MovCap MC ON M.nMovNro = MC.nMovNro" _
          & "   INNER JOIN MovCapDet MCD ON MC.nMovNro = MCD.nMovNro And MC.cOpeCod  = MCD.cOpeCod And MC.cCtaCod  = MCD.cCtaCod and not nConceptoCod in (6,7) " _
          & "   LEFT  JOIN MovCMAC MCMAC ON M.nMovNro = MCMAC.nMovNro" _
          & "   LEFT  JOIN Captaciones C ON MC.cCtaCod = C.cCtaCod" _
          & "   LEFT  JOIN ColocRecup CJ ON MC.cCtaCod = CJ.cCtaCod" _
          & "   LEFT  JOIN MovDoc MD ON M.nMovNro = MD.nMovNro" _
          & "   Where (m.nMovFlag = 0) And MCD.nMonto <> 0  And M.cMovNro Like '" & Format(pdFechaAsiento, oFunV.gsFormatoMovFecha) & "%'" _
          & "       AND substring(M.cOpeCod,1,3) Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoNoAsientoCab3 & ")" _
          & "       And substring(M.cOpeCod,1,4) Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoNoAsientoCab4 & ")" _
          & "       And MC.cOpeCod NOT IN (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoSiColocCargoCuenta & ")" _
          & "       And M.cOpeCod Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoNoAsiento & ")" _
          & "       AND substring(M.cOpeCod,1,2) IN (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoSiAsiento & ") "
    
    sql = sql & " Union SELECT 0 as nCredEstado, M.nMovNro, M.cMovNro, dbo.FechaHoraMov(M.cMovNro) dFecTran, '108' + Substring(M.cMovNro,18,2) + '000' + Ltrim(Rtrim(str(nMoneda))) cCtaCod, M.cOpeCod,  0 Concepto, round(abs(MC.nMovImporte),2) As SumaMonto,  SubString(m.cMovNro,18,2) cCodAge, MC.cNroDoc cDocNro, 0 nNroCuota, MC.nMovImporte  nSaldCnt, Null As LineaC, NULL AS cRFA" _
          & " FROM Mov M" _
          & " INNER JOIN MovOpeVarias MC ON M.nMovNro = MC.nMovNro" _
          & " Where (m.nMovFlag = 0) And MC.nMovImporte <> 0  And M.cMovNro Like '" & Format(pdFechaAsiento, oFunV.gsFormatoMovFecha) & "%' AND" _
          & " (substring(M.cOpeCod,1,3) Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = 5)" _
          & " And substring(M.cOpeCod,1,4) Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = 6)" _
          & " And M.cOpeCod Not In            (Select cOpeCod From OpeTpoSinAsiento Where nProceso = 3)" _
          & " AND substring(M.cOpeCod,1,2) IN     (Select cOpeCod From OpeTpoSinAsiento Where nProceso = 4)" _
          & " or M.cOpeCod In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = 0)) " _
    
    sql = sql & " Union SELECT 0 as nCredEstado, M.nMovNro, M.cMovNro, dbo.FechaHoraMov(M.cMovNro) dFecTran, '108' + Substring(M.cMovNro,18,2) + '000' + Ltrim(Rtrim(str(nMoneda))) cCtaCod, M.cOpeCod,  0 Concepto, round(abs(MC.nMovImporte),2) As SumaMonto,  SubString(m.cMovNro,18,2) cCodAge, cDocNro = '', 0 nNroCuota, MC.nMovImporte  nSaldCnt, Null As LineaC, NULL AS cRFA" _
          & " FROM Mov M INNER JOIN MovHabilitacion MC ON M.nMovNro = MC.nMovNro" _
          & " Where (m.nMovFlag = 0) And MC.nMovImporte <> 0  And M.cMovNro Like '" & Format(pdFechaAsiento, oFunV.gsFormatoMovFecha) & "%' AND" _
          & " substring(M.cOpeCod,1,3) Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = 5)" _
          & " And substring(M.cOpeCod,1,4) Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = 6)" _
          & " And M.cOpeCod Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = 3)"
    
    sql = sql & " Union SELECT 0 as nCredEstado, M.nMovNro, M.cMovNro, dbo.FechaHoraMov(M.cMovNro) dFecTran, '108' + Substring(M.cMovNro,18,2) + '000' + Ltrim(Rtrim(str(nMoneda))) cCtaCod, M.cOpeCod,  0 Concepto, round(abs(MC.nMovImporte),2) As SumaMonto,  SubString(M.cMovNro,18,2) cCodAge, cDocNro = '', 0 nNroCuota, MC.nMovImporte  nSaldCnt, Null As LineaC, NULL AS cRFA" _
          & " FROM Mov M JOIN MovRef MR JOIN Mov M1 INNER JOIN MovHabilitacion MC ON M1.nMovNro = MC.nMovNro ON MR.nMovNroRef = M1.nMovNro ON M.nMovNro = MR.nMovNro " _
          & " Where (M.nMovFlag = 0) And MC.nMovImporte <> 0  And M.cMovNro Like '" & Format(pdFechaAsiento, oFunV.gsFormatoMovFecha) & "%' AND" _
          & " substring(M.cOpeCod,1,3) Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = 5) " _
          & " And substring(M.cOpeCod,1,4) Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = 6) " _
          & " And M.cOpeCod Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = 3)"
    
    sql = sql & " Union SELECT 0 as nCredEstado, M.nMovNro, M.cMovNro, dbo.FechaHoraMov(M.cMovNro) dFecTran, '108' + Substring(M.cMovNro,18,2) + '000' + Ltrim(Rtrim(str(nMoneda))) cCtaCod, M.cOpeCod,  MSD.nPrdConcepto Concepto, round(abs(MSD.nMonto),2) As SumaMonto,  SubString(M.cMovNro,18,2) cCodAge, cDocNro = '', 0 nNroCuota, MSD.nMonto nSaldCnt, Null As LineaC, NULL AS cRFA" _
          & " FROM Mov M JOIN MovServicios MS JOIN MovServiciosDet MSD ON MS.nMovNro = MSD.nMovNro And MS.cNumRecibo = MSD.cNumRecibo ON M.nMovNro = MS.nMovNro " _
          & " Where (M.nMovFlag = 0) And MS.nMonto <> 0  And M.cMovNro Like '" & Format(pdFechaAsiento, oFunV.gsFormatoMovFecha) & "%' "
    
    'REMESAS A CAJA GENERAL -- SIMON --
'    sql = sql & " Union SELECT M.nMovNro, M.cMovNro, dbo.FechaHoraMov(M.cMovNro) dFecTran, '112' + Substring(M.cMovNro,18,2) + '000' + LEFT(cEfectivoCod,1) cCtaCod, M.cOpeCod,  0 Concepto, MC.nMonto As SumaMonto,  SubString(m.cMovNro,18,2) cCodAge, cDocNro = '', 0 nNroCuota, MC.nMonto nSaldCnt, Null As LineaC" _
'          & " FROM Mov M INNER JOIN MovUserEfectivo MC ON M.nMovNro = MC.nMovNro" _
'          & " Where (m.nMovFlag = 0) And MC.nMonto <> 0  And M.cMovNro Like '" & Format(pdFechaAsiento, oFunV.gsFormatoMovFecha) & "%' AND" _
'          & " substring(M.cOpeCod,1,3) Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = 5)" _
'          & " And substring(M.cOpeCod,1,4) Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = 6)" _
'          & " And M.cOpeCod Not In (Select cOpeCod From OpeTpoSinAsiento Where nProceso = 3) " _
'          & " And M.cOpeCod Not IN ('901007','901016')"

    sql = sql & " ORDER BY MC.cCtaCod "
    Set rs = oCon.CargaRecordSet(sql)
    If Not (rs.EOF And rs.BOF) Then
        Set CargaTranDiaria = rs
    Else
        Set CargaTranDiaria = Nothing
    End If
    oCon.CierraConexion
    Set oFunV = Nothing
End Function

Public Function ActualizaCapitalizacionAhorrosCTS(ByVal pnPaso As Integer, ByVal gAsientoProcesoCapAho As String, ByVal gAsientoProcesoCapCTS As String, ByVal gsCodPersCMACT As String, ByVal pdFechaAsiento As Date) As ADODB.Recordset
    Dim ssql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim oFunV As New COMFunciones.FCOMVarPublicas
    oCon.AbreConexion
            If pnPaso = 1 Then
                ssql = " Select   A.cCtaCod, SUBSTRING(A.cCtaCod,4,2) cAgencia, SUBSTRING(A.cCtaCod,9,1) cMoneda, M.cOpeCod, A.nPersoneria, " _
                       & "          OC.cCtaContCod, OC.cOpeCtaDH, SUM(Abs(MC.nMonto)) Monto" _
                       & " FROM OpectaNeg OC" _
                       & "          INNER JOIN Mov M" _
                       & "          INNER JOIN MovCap MC ON M.nMovNro = MC.nMovNro" _
                       & "          INNER JOIN Captaciones A ON MC.cCtaCod = A.cCtaCod ON OC.cOpeCod = M.cOpeCod AND OC.nPersoneria = A.nPersoneria" _
                       & " Where M.cOpeCod IN (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoCapAho & ")" _
                       & "       And (M.nMovFlag = 0) " _
                       & "       AND A.nPersoneria NOT IN (4,5,7,8) " _
                       & "       And M.cMovNro Like '" & Format(pdFechaAsiento, oFunV.gsFormatoMovFecha) & "%'" _
                       & "  Group by A.cCtaCod,SUBSTRING(A.cCtaCod,4,2), substring(A.cCtaCod,9,1), M.cOpeCod, A.nPersoneria, OC.cCtaContCod, OC.cOpeCtaDH" _
                       & " Order by A.cCtaCod,cAgencia, cMoneda, M.cOpeCod, A.nPersoneria, OC.cOpeCtaDH"

            ElseIf pnPaso = 2 Then
                'Personeria 4,5,7,8
                ssql = " Select   A.cCtaCod,SUBSTRING(A.cCtaCod,4,2) cAgencia, SUBSTRING(A.cCtaCod,9,1) cMoneda, M.cOpeCod, A.nPersoneria, I.cSubCtaContCod, " _
                       & "          OC.cCtaContCod, OC.cOpeCtaDH, SUM(Abs(MC.nMonto)) Monto" _
                       & " FROM OpectaNeg OC " _
                       & "      JOIN Mov M JOIN MovCap MC " _
                       & "      JOIN Captaciones A " _
                       & "      JOIN ProductoPersona PP " _
                       & "      JOIN InstitucionFinanc I ON PP.cPersCod = I.cPersCod ON A.cCtaCod = PP.cCtaCod ON " _
                       & "      MC.cCtaCod = A.cCtaCod ON M.nMovNro = MC.nMovNro ON OC.cOpeCod = M.cOpeCod AND OC.nPersoneria = A.nPersoneria " _
                       & " Where M.cOpeCod IN (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoCapAho & ")" _
                       & "      And (M.nMovFlag = 0) And PP.nPrdPersRelac = 10 " _
                       & "      And M.cMovNro Like '" & Format(pdFechaAsiento, oFunV.gsFormatoMovFecha) & "%'" _
                       & " Group by A.cCtaCod,SUBSTRING(A.cCtaCod,4,2), substring(A.cCtaCod,9,1), M.cOpeCod, A.nPersoneria, I.cSubCtaContCod, OC.cCtaContCod, OC.cOpeCtaDH" _
                       & " Order by A.cCtaCod,cAgencia, cMoneda, M.cOpeCod, A.nPersoneria, I.cSubCtaContCod, OC.cOpeCtaDH"
                        
            ElseIf pnPaso = 3 Then
                'CTS
                ssql = "Select    T1.cCtaCod, T1.cAgencia, T1.cTipoCliente, T1.cMoneda, T1.cCtaContCod, T1.cOpeCtaDH, Sum(T1.Monto) Monto, T1.cOpeCod " _
                    & "From " _
                    & "(    Select  CaptacCTS.cCtaCod, cTipoCliente = CASE WHEN LTRIM(RTRIM(CaptacCTS.cCodInst)) = '" & gsCodPersCMACT & "' " _
                    & "             THEN '02' Else '01' END, " _
                    & "             SUBSTRING(CaptacCTS.cCtaCod,9,1) cMoneda, SUBSTRING(CaptacCTS.cCtaCod,4,2) cAgencia, " _
                    & "             LTRIM(RTRIM(OC.cCtaContCod)) cCtaContCod, OC.cOpeCtaDH, Abs(MC.nMonto) Monto, T.cOpeCod " _
                    & "     FROM    Mov T " _
                    & "             INNER JOIN  MovCap MC " _
                    & "             INNER JOIN  MovCapDet MCD " _
                    & "             INNER JOIN  OpeCtaNeg OC ON MCD.cOpeCod = OC.cOpeCod And MCD.nConceptoCod = nConcepto ON MC.nMovNro = MCD.nMovNro And " _
                    & "                         MC.cOpeCod = MCD.cOpeCod And MC.cCtaCod = MCD.cCtaCod ON T.nMovNro = MC.nMovNro " _
                    & "             INNER JOIN CaptacCTS ON MC.cCtaCod = CaptacCTS.cCtaCod " _
                    & "     WHERE   T.cMovnro like '" & Format(pdFechaAsiento, oFunV.gsFormatoMovFecha) & "%' " _
                    & "             And T.cOpeCod IN (Select cOpeCod From OpeTpoSinAsiento " _
                    & "                       Where nProceso = " & gAsientoProcesoCapCTS & ") AND (T.nMovFlag = 0) ) T1 " _
                    & "GROUP BY    T1.cCtaCod, cAgencia, cTipoCliente, cMoneda, cCtaContCod, cOpeCtaDH,cOpeCod Order by cAgencia, cTipoCliente, cMoneda, cCtaContCod, cOpeCtaDH"
             End If
                    
    Set rs = oCon.CargaRecordSet(ssql)
    If Not (rs.EOF And rs.BOF) Then
        Set ActualizaCapitalizacionAhorrosCTS = rs
    Else
        Set ActualizaCapitalizacionAhorrosCTS = Nothing
    End If
    oCon.CierraConexion
    Set oFunV = Nothing
End Function
    
Public Function CargaOpeCuenta(ByVal pscOpeCod As String, ByVal psConcepto As String, ByVal psSql As String) As ADODB.Recordset
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
     
    oCon.AbreConexion
    sql = " SELECT cCtaContCod , cOpeCtaDH FROM OpeCtaNeg  " & _
          " WHERE cOpeCod = '" & pscOpeCod & "' And nConcepto = " & psConcepto & " " & psSql
          
    Set rs = oCon.CargaRecordSet(sql)
    If Not (rs.EOF And rs.BOF) Then
        Set CargaOpeCuenta = rs
    Else
        Set CargaOpeCuenta = Nothing
    End If
    oCon.CierraConexion
End Function

Public Function ClienteTipoPersCol(ByVal nMovNro As Long) As String
    Dim RegPer As ADODB.Recordset
    Dim tmpSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    tmpSql = "SELECT C.nPersoneria FROM Captaciones C JOIN MovCap MC JOIN Mov M ON " _
        & "MC.nMovNro = M.nMovNro ON C.cCtaCod = MC.cCtaCod WHERE M.nMovNro = " & nMovNro
    
    
    Set RegPer = New ADODB.Recordset
    RegPer.CursorLocation = adUseClient
    Set RegPer = oCon.CargaRecordSet(tmpSql)
    Set RegPer.ActiveConnection = Nothing
    If (RegPer.BOF Or RegPer.EOF) Then
        ClienteTipoPersCol = ""
    Else
        ClienteTipoPersCol = Trim(RegPer!nPersoneria)
    End If
    RegPer.Close
    Set RegPer = Nothing
End Function

Public Function VarAsientoPrestamosAdm(ByVal sCuenta As String) As String
    Dim RegCta As ADODB.Recordset
    Dim tmpSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    Set RegCta = New ADODB.Recordset
    oCon.AbreConexion
    tmpSql = "Select C.* from ColocGarantia C JOIN Garantias G ON C.cNumGarant = G.cNumGarant " _
        & "Where C.cCtaCod IN ('" & sCuenta & "') And C.nEstado = 1 And G.nEstado IN (3,6) " _
        & "And G.cTpoDoc = '121'"
    
    Set RegCta = oCon.CargaRecordSet(tmpSql)
    If (RegCta.BOF Or RegCta.EOF) Then
        VarAsientoPrestamosAdm = "02"
    Else
        VarAsientoPrestamosAdm = "01"
    End If
    RegCta.Close
    Set RegCta = Nothing
End Function

Public Function ObtenerCtaNeg(ByVal pscOpeCod As String) As ADODB.Recordset
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    sql = " SELECT cCtaContCod , cOpeCtaDH  FROM OpeCtaNeg  " _
            & " WHERE cOpeCod = '" & pscOpeCod & "'"
    Set rs = oCon.CargaRecordSet(sql)
    
    If (rs.BOF Or rs.EOF) Then
        Set ObtenerCtaNeg = Nothing
    Else
        Set ObtenerCtaNeg = rs
    End If
    oCon.CierraConexion
End Function


Public Function ObtenerSobrantesFaltantes(ByVal gAsientoProcesoSiAsientoSofFal As String, ByVal pdFechaAsiento As Date) As ADODB.Recordset
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim oFunV As New COMFunciones.FCOMVarPublicas
    oCon.AbreConexion
    
    sql = " Select SUBSTRING(M.cMovNro,18,2) Agencia, ABS(O.nMovImporte) As Monto, M.cOpeCod as cCodOpe, O.nMoneda FROM Mov M " _
            & " Inner Join MovOpeVarias O On M.nMovNro = O.nMovNro " _
            & " Where M.cMovNro like '" & Format(pdFechaAsiento, oFunV.gsFormatoMovFecha) & "%' And M.cOpeCod " _
            & " IN (Select cOpeCod From OpeTpoSinAsiento Where nProceso = " & gAsientoProcesoSiAsientoSofFal & ")  and (M.nMovFlag = 0) AND nMovImporte <> 0"
            
    Set rs = oCon.CargaRecordSet(sql)
    If (rs.BOF Or rs.EOF) Then
        Set ObtenerSobrantesFaltantes = Nothing
    Else
        Set ObtenerSobrantesFaltantes = rs
    End If
    oCon.CierraConexion
    
End Function


Public Function CierreRealizado(ByVal pdfecsis As Date, Optional pnTipo As Integer = 1) As Boolean
Dim rsVarSis As ADODB.Recordset
Dim pdCieDia As Date
Dim pdCieMes As Date
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
Set rsVarSis = New ADODB.Recordset
Dim sql As String

oCon.AbreConexion

    rsVarSis.CursorLocation = adUseClient
    sql = "select nConsSisCod,nConsSisDesc, nConsSisValor From ConstSistema where nConsSisCod In (" & ConstSistemas.gConstSistCierreSistema & "," & ConstSistemas.gConstSistCierreMesNegocio & ")"
    Set rsVarSis = oCon.CargaRecordSet(sql)
    Set rsVarSis.ActiveConnection = Nothing
    Do While Not rsVarSis.EOF
        If Trim(rsVarSis!nConsSisCod) = ConstSistemas.gConstSistCierreSistema Then
            pdCieDia = CDate(Trim(rsVarSis!nConsSisValor))
        ElseIf Trim(rsVarSis!nConsSisCod) = ConstSistemas.gConstSistCierreMesNegocio Then
            pdCieMes = CDate(Trim(rsVarSis!nConsSisValor))
        End If
        rsVarSis.MoveNext
    Loop
    rsVarSis.Close
    Set rsVarSis = Nothing
    If pnTipo = 1 Then
        CierreRealizado = IIf(pdCieDia = pdfecsis, True, False)
    ElseIf pnTipo = 2 Then
        CierreRealizado = IIf(pdCieMes = pdfecsis, True, False)
    End If

End Function

Public Function ObtenerDiasSinMov(ByVal pdHoraGrab As String) As ADODB.Recordset
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim oFunF As New COMFunciones.FCOMVarPublicas
    oCon.AbreConexion
    
    sql = "Select cMovNro From Mov M Where M.cMovNro Like '" & Format(pdHoraGrab, oFunF.gsFormatoMovFecha) & "%' AND M.nMovFlag = " & COMDConstantes.gMovFlagVigente
    
    Set rs = oCon.CargaRecordSet(sql)
    If (rs.BOF Or rs.EOF) Then
        Set ObtenerDiasSinMov = Nothing
    Else
        Set ObtenerDiasSinMov = rs
    End If
    
    oCon.CierraConexion
    Set oFunF = Nothing
End Function
 Public Function GetTipCambio(pdFecha As Date) As Boolean
Dim oDGeneral As COMDConstSistema.NCOMTipoCambio
Set oDGeneral = New COMDConstSistema.NCOMTipoCambio
 GetTipCambio = True
 gnTipCambio = 0
 gnTipCambioV = 0
 gnTipCambioC = 0

 gnTipCambio = oDGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
 gnTipCambioV = oDGeneral.EmiteTipoCambio(pdFecha, TCVenta)
 gnTipCambioC = oDGeneral.EmiteTipoCambio(pdFecha, TCCompra)

If gnTipCambio = 0 Then
    GetTipCambio = False
End If
End Function

Public Function ExisCtaCnt(ByVal pCtaCnt As String, Optional ByVal pNuePlan As Boolean = False) As Boolean
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    tmpSql = "SELECT cCtaContDesc FROM CtaCont WHERE cCtaContCod LIKE '" & pCtaCnt & "%'"
    Set tmpReg = oCon.CargaRecordSet(tmpSql)
    If (tmpReg.BOF Or tmpReg.EOF) Then
        ExisCtaCnt = False
    Else
        ExisCtaCnt = IIf(tmpReg.RecordCount > 1, False, True)
    End If
    tmpReg.Close
    Set tmpReg = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function VarAG(ByVal pCuentaProd As String, ByVal pCtaCont As String) As String
    Dim lsAge As String
    '******** ica parche en caja agencias EJRS 13/10/2004 *****************************
    If Left(pCtaCont, 2) = "11" Then
        lsAge = Mid(pCuentaProd, 4, 2)
    Else
        lsAge = Mid(pCuentaProd, 4, 2)
 
    End If
    '**** CULMINACION DE PARCHES JEAN 13/10/2004 ***********+
    VarAG = lsAge
End Function

Public Function ClienteTipoPers(ByVal pCuenta As String) As String
    Dim RegPer As ADODB.Recordset
    Dim tmpSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    
    Select Case Mid(pCuenta, 6, 3)
        Case Producto.gCapAhorros, Producto.gCapCTS, Producto.gCapPlazoFijo
            tmpSql = "SELECT nPersoneria FROM Captaciones WHERE cCtaCod = '" & pCuenta & "'"
        Case Else
            tmpSql = " SELECT Max(nPersPersoneria) nPersoneria FROM Persona AS P, ProductoPersona AS PC " & _
                     " WHERE P.cPersCod = PC.cPersCOd AND PC.cCtaCod = '" & pCuenta & "'" & _
                     " AND PC.nPrdPersRelac = " & ColocRelacPers.gColRelPersTitular
    End Select
    
    Set RegPer = New ADODB.Recordset
    RegPer.CursorLocation = adUseClient
    Set RegPer = oCon.CargaRecordSet(tmpSql)
    Set RegPer.ActiveConnection = Nothing
    If (RegPer.BOF Or RegPer.EOF) Then
        ClienteTipoPers = ""
    Else
        ClienteTipoPers = Trim(RegPer!nPersoneria)
    End If
    RegPer.Close
    Set RegPer = Nothing
End Function

Private Function VerificaCuenta(sCtaCod As String, ByRef sObjetoCod As String, ByRef sIFTpo As String, ByRef SCtaIfCod As String) As String
Dim lsCta As String
Dim CtaCor As String
Dim SubCta As String
Dim rs     As New ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
Dim ssql As String
oCon.AbreConexion

lsCta = sCtaCod
sObjetoCod = ""
ssql = "SELECT cCtaContCod, MAX(cPersCod+cIFTpo+cCtaIFCod) as cObjetoCod FROM CtaIFFiltro WHERE cCtaContCod + cCtaIFSubCta = '" & sCtaCod & "' GROUP BY cCtaContCod "
If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
Set rs = oCon.CargaRecordSet(ssql)
If Not rs.EOF Then
   sObjetoCod = Left(rs!cObjetoCod, 13)
   sIFTpo = Mid(rs!cObjetoCod, 14, 2)
   SCtaIfCod = Mid(rs!cObjetoCod, 16, 7)
End If
rs.Close: Set rs = Nothing
VerificaCuenta = lsCta
End Function

Public Function VarBC(ByVal nMovNro As Long, ByVal sCtaConCod As String) As String
Dim RegCta As ADODB.Recordset
Dim tmpSql As String, sCtaCod As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
    
oCon.AbreConexion

sCtaCod = Mid(sCtaConCod, 1, InStr(1, sCtaConCod, "BC", vbTextCompare) - 1)

tmpSql = "SELECT ISNULL(cCtaIFSubCta,'') Campo FROM Mov M JOIN MovOpeVarias V JOIN CtaIFFiltro F ON " _
    & "SUBSTRING(V.cReferencia,4,13) = F.cPersCod AND LEFT(V.cReferencia,2) = F.cIFTpo " _
    & "AND SUBSTRING(V.cReferencia,18,7) = F.cCtaIfCod ON M.nMovNro = V.nMovNro " _
    & "WHERE M.nMovNro = " & nMovNro

Set RegCta = oCon.CargaRecordSet(tmpSql)
    If (RegCta.BOF Or RegCta.EOF) Then
        VarBC = ""
    Else
        VarBC = Trim(RegCta!Campo)
    End If
    RegCta.Close
    Set RegCta = Nothing
End Function

Public Function VarPL(ByVal pCuenta As String, ByVal pCodLinC As String, ByVal pCodLinCJ As String) As String
    Dim vCodLinea As String
    
    'If Left(pCuenta, 2) = Right(gsCodAge, 2) Then  ' Para Judicial
       vCodLinea = pCodLinC
    'Else
    '   vCodLinea = pCodLinCJ
    'End If
    VarPL = Mid(Trim(vCodLinea), 6, 1)
End Function

Public Function VarFF(ByVal pCuenta As String, ByVal pCodLinC As String, ByVal pCodLinCJ As String) As String
    Dim RegCta As New ADODB.Recordset
    Dim tmpSql1 As String, sServConsol As String
    Dim oCon As COMConecta.DCOMConecta
    
    Dim loConstSist As COMDConstSistema.NCOMConstSistema
    Set loConstSist = New COMDConstSistema.NCOMConstSistema
    sServConsol = loConstSist.LeeConstSistema(gConstSistServCentralRiesgos)
    Set loConstSist = Nothing

    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion
    '** EQUIVALENTE FUENTE FINANCIAMIENTO
    tmpSql1 = "Select cCtaCont FROM " & sServConsol & "ColocLineaCreditoEquiv WHERE cLineaCred = '" & Mid(pCodLinC, 1, 4) & "' "
    Set RegCta = oCon.CargaRecordSet(tmpSql1)
    If (RegCta.BOF Or RegCta.EOF) Then
        VarFF = ""
    Else
       VarFF = RegCta!cCtaCont
    End If
    
End Function

Public Function VarPD(ByVal pCuenta As String, ByVal pCodLinC As String, ByVal pCodLinCJ As String) As String
    Dim vCodLinea As String
    Dim vcad As String
    vCodLinea = pCodLinC
        
    vcad = Mid(Trim(pCuenta), 6, 3)
    If (vcad = "101" Or vcad = "201") And Mid(Trim(vCodLinea), 6, 1) = "2" Then
        VarPD = "05"
    ElseIf vcad = "101" Or vcad = "201" Then
        VarPD = "01"
    ElseIf vcad = "301" Or vcad = "302" Or vcad = "303" Or vcad = "304" Or vcad = "320" Then
        VarPD = "02"
    ElseIf vcad = "102" Or vcad = "202" Then
        VarPD = "03"
    ElseIf vcad = "401" Or vcad = "423" Or vcad = "402" Then
        VarPD = "02"
    Else
        VarPD = ""
    End If
End Function


Public Function Billetaje(ByVal pFecha As Date, ByVal pMoneda As String, Optional psAgeCod As String = "") As Currency
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    Dim oFunV As New COMFunciones.FCOMVarPublicas
    
    oCon.AbreConexion
    
tmpSql = "Select ISNULL(SUM(E.nMonto),0) Campo From Mov M INNER JOIN MovUserEfectivo E ON M.nMovNro = E.nMovNro " _
    & "Where M.nMovFlag NOT IN (2,1)  And LEFT(M.cMovNro,8) IN (Select MAX(LEFT(M1.cMovNro,8)) From Mov M1 " _
    & "JOIN MovUserEfectivo E1 ON M1.nMovNro = E1.nMovNro Where M1.nMovFlag NOT IN (2,1) And E1.cEfectivoCod LIKE '" & pMoneda & "%' And " _
    & "LEFT(M1.cMovNro,8) <= '" & Format(pFecha, oFunV.gsFormatoMovFecha) & "' AND M.cOpeCod in ('901007', '901016')) AND M.cOpeCod in ('901007', '901016') And E.cEfectivoCod LIKE '" & pMoneda & "%'"
If Not psAgeCod = "" Then
    tmpSql = tmpSql & " and substring(cmovnro,18,2) = '" & psAgeCod & "'"
End If

    Set tmpReg = oCon.CargaRecordSet(tmpSql)
    
    If (tmpReg.BOF Or tmpReg.EOF) Then
        Billetaje = 0
    Else
        Billetaje = IIf(IsNull(tmpReg!Campo), 0, tmpReg!Campo)
    End If
    tmpReg.Close
    Set tmpReg = Nothing
    Set oCon = Nothing
End Function


Public Function EstadCred(ByVal pFecha As Date, ByVal pMoneda As String, ByVal pOption As Integer, Optional psAgeCod As String = "") As Currency
Dim tmpReg As ADODB.Recordset
Dim tmpSql As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

oCon.AbreConexion

If pOption = 1 Then
    tmpSql = "SELECT (ISNULL(SUM(CASE when datediff(dd,dEstad,'" & Format(DateAdd("d", -1, pFecha), "mm/dd/yyyy") & "') = 0 then nSaldoCap END),0)) - " & _
        " (ISNULL(SUM(CASE when datediff(dd,dEstad,'" & Format(pFecha, "mm/dd/yyyy") & "') = 0 then nSaldoCap END),0) ) AS Campo " & _
        " FROM ColocEstadDiaCred WHERE substring(cLineaCred,5,1) = '" & pMoneda & "'" & _
        IIf(psAgeCod = "", "", " and cCodAge = '" & psAgeCod & "'")
Else
    tmpSql = "SELECT (ISNULL(SUM(CASE When datediff(dd,dEstad,'" & Format(DateAdd("d", -1, pFecha), "mm/dd/yyyy") & "') = 0 THEN nCapVig END),0)) - " & _
        " (ISNULL(SUM(CASE when datediff(dd,dEstad,'" & Format(pFecha, "mm/dd/yyyy") & "') = 0 then nCapVig END),0)) AS Campo " & _
        " FROM ColocEstadDiaPrenda " & _
        IIf(psAgeCod = "", "", " WHERE right(cCodAge,2) = '" & psAgeCod & "'")
End If
Set tmpReg = New ADODB.Recordset
tmpReg.CursorLocation = adUseClient
Set tmpReg = oCon.CargaRecordSet(tmpSql)
Set tmpReg.ActiveConnection = Nothing
If (tmpReg.BOF Or tmpReg.EOF) Then
    EstadCred = 0
Else
    EstadCred = IIf(IsNull(tmpReg!Campo), 0, tmpReg!Campo)
End If
tmpReg.Close
Set tmpReg = Nothing
Set oCon = Nothing
End Function

Public Function EstadAho(ByVal pFecha As Date, ByVal pMoneda As String, ByVal pOption As Integer, Optional psAgeCod As String = "") As Currency
Dim tmpReg As ADODB.Recordset
Dim tmpSql As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
Dim oFunV As New COMFunciones.FCOMVarPublicas
oCon.AbreConexion

If pOption = 1 Then
    tmpSql = " SELECT (ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(DateAdd("d", -1, pFecha), oFunV.gsFormatoFecha) & "') = 0 then nSaldo End)),0)" _
           & "         -  ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(pFecha, oFunV.gsFormatoFecha) & "') = 0 then nSaldo End)),0)) AS Campo" _
           & " FROM CapEstadSaldo WHERE nMoneda = '" & pMoneda & "' And nProducto = " & Producto.gCapAhorros & _
        IIf(psAgeCod = "", "", " and cCodAge = '" & psAgeCod & "'")
ElseIf pOption = 2 Then
    tmpSql = " SELECT    (ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(DateAdd("d", -1, pFecha), oFunV.gsFormatoFecha) & "') = 0 then nSaldo End)),0)" _
           & "         -  ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(pFecha, oFunV.gsFormatoFecha) & "') = 0 then nSaldo End)),0)) AS Campo" _
           & " FROM CapEstadSaldo WHERE nMoneda = '" & pMoneda & "' And nProducto = " & Producto.gCapPlazoFijo & _
        IIf(psAgeCod = "", "", " and cCodAge = '" & psAgeCod & "'")
Else
    tmpSql = " SELECT    (ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(DateAdd("d", -1, pFecha), oFunV.gsFormatoFecha) & "') = 0 then nSaldo End)),0)" _
           & "         -  ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(pFecha, oFunV.gsFormatoFecha) & "') = 0 then nSaldo End)),0)) AS Campo" _
           & " FROM CapEstadSaldo WHERE nMoneda = '" & pMoneda & "' And nProducto = " & Producto.gCapCTS & _
        IIf(psAgeCod = "", "", " and cCodAge = '" & psAgeCod & "'")
End If
Set tmpReg = New ADODB.Recordset
tmpReg.CursorLocation = adUseClient
Set tmpReg = oCon.CargaRecordSet(tmpSql)
Set tmpReg.ActiveConnection = Nothing
If (tmpReg.BOF Or tmpReg.EOF) Then
    EstadAho = 0
Else
    EstadAho = IIf(IsNull(tmpReg!Campo), 0, tmpReg!Campo)
End If
tmpReg.Close
Set tmpReg = Nothing
End Function



Public Function Repetido(ByVal pCadena As String, ByVal pBusca As String) As Integer
    Dim vPos As Integer
    Repetido = 0
    Do While Len(pCadena) > 0
        vPos = InStr(1, pCadena, pBusca, vbTextCompare)
        If vPos > 0 Then
            pCadena = Mid(pCadena, vPos + 1)
            Repetido = Repetido + 1
        Else
            Exit Do
        End If
    Loop
End Function

Public Function CuentaNombre(ByVal pCodCta As String, Optional ByVal pNuePlan As Boolean = False) As String
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    Dim pcodcta2 As String, pcodcta3 As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    pCodCta = Trim(pCodCta)
    If Len(pCodCta) > 4 Then pcodcta2 = Left(pCodCta, Len(pCodCta) - 2)
    If Len(pCodCta) > 6 Then pcodcta3 = Left(pCodCta, Len(pCodCta) - 4)

    oCon.AbreConexion
    
    tmpSql = "SELECT cCtaContDesc FROM  CtaCont WHERE cCtaContCod IN ('" & Left(pCodCta, 4) & "','" & _
        pCodCta & "','" & pcodcta2 & "','" & pcodcta3 & "') order by cCtaContCod"
    Set tmpReg = oCon.CargaRecordSet(tmpSql)
    If (tmpReg.BOF Or tmpReg.EOF) Then
        CuentaNombre = ""
    Else
        Do While Not tmpReg.EOF
            CuentaNombre = Trim(CuentaNombre) & " " & Trim(tmpReg!cCtaContDesc)
            tmpReg.MoveNext
        Loop
    End If
    tmpReg.Close
    Set tmpReg = Nothing
End Function

'**************************** GENERACION DEL ASIENTO ***************************************************
'*****************************************************************************************************

Public Function GenerarAsiento(ByVal dFecha As Date, ByVal pnaho As Integer, ByVal pncred As Integer, ByVal pnpig As Integer, ByVal pnMonSol As Integer, ByVal pnMonDol As Integer, _
                               ByVal pdfecsis As Date, ByVal psNomCmac As String, ByVal psCodAge As String, ByVal psNomAge As String, ByVal pscodcmac As String, ByVal pscoduser As String, ByVal gsCodPersCMACT As String, ByVal pAsiDia As String, _
                               ByRef psmensaje As String, ByRef psvobs As String, ByRef psvnoctacnt As String, ByRef pbpasidia As String, ByRef pbcierrerealizado As String, ByRef psmenError As String) As String
   
    Dim oFunI As New COMFunciones.FCOMImpresion
    Dim oFunV As New COMFunciones.FCOMVarPublicas
    Dim oPers As COMDPersona.DCOMInstFinac
    
    Dim lsCVMEGanacia As String
    Dim lsCVMEPerdida As String
    Dim lsCajaSoles As String
    Dim lsCajaDolares As String
    Dim lsCajaAgenciaSoles As String
    Dim lsCajaAgenciaDolares As String
    
    Dim pPrevioMax As Double
    Dim pLineasMax As Double
    Dim pHojaFiMax As Integer
    Dim ssql As String
    Dim vRTFImp As String
    Dim dHoraGrab As Date
    Dim gdHoraGrab As String
    Dim gbAsientoDN As Boolean
    
    Dim lTransActiva As Boolean
    Dim sObjetoCod  As String
    Dim sIFTpo As String
    Dim SCtaIfCod As String

    Dim lsmenError As String
    Dim gsCtaCodFoncodes As String
    
    DoEvents

    gsCtaCodFoncodes = GetAsientoParametro(1)
    lsCVMEGanacia = GetAsientoParametro(2)
    lsCVMEPerdida = GetAsientoParametro(3)
    lsCajaSoles = GetAsientoParametro(4)
    lsCajaDolares = GetAsientoParametro(5)
    lsCajaAgenciaSoles = GetAsientoParametro(6)
    lsCajaAgenciaDolares = GetAsientoParametro(7)
    
    Dim RegTran As ADODB.Recordset
    Dim RegOpeCta As ADODB.Recordset
    Dim RegTmp As ADODB.Recordset
    Dim tmpSql As String
    Dim vMoneda As COMDConstantes.Moneda
    Dim vAgencia As String
    Dim vCodConta As String
    Dim vNoCtaCnt As String
    Dim vParche As String
    Dim vCont As Long, vCtaCnt As Long
    Dim vEst As String
    Dim vcad As String
    Dim x As Long
    Dim TCF As Currency, TCC As Currency, TCV As Currency
    Dim sTmp1 As String, sTmp3 As String
    Dim vespacio As Long, vLenNomb As Long
    Dim Arreglo() As Currency
    
    Dim gAsientoProcesoCapAho As Long
    Dim gAsientoProcesoCapCTS As Long
    Dim gAsientoProcesoNoAsiento As Long
    Dim gAsientoProcesoSiAsiento As Long
    Dim gAsientoProcesoNoAsientoCab3 As Long
    Dim gAsientoProcesoNoAsientoCab4 As Long
    Dim gAsientoProcesoSiAsientoCVME As Long
    Dim gAsientoProcesoSiAsientoSofFal As Long
    Dim gAsientoProcesoSiColocCargoCuenta As Long
    Dim gAsientoProcesoCreditoSinAsiento As Long
    
    gAsientoProcesoCapAho = 1
    gAsientoProcesoCapCTS = 2
    gAsientoProcesoNoAsiento = 3
    gAsientoProcesoSiAsiento = 4
    gAsientoProcesoNoAsientoCab3 = 5
    gAsientoProcesoNoAsientoCab4 = 6
    gAsientoProcesoSiAsientoCVME = 7
    gAsientoProcesoSiAsientoSofFal = 8
    gAsientoProcesoSiColocCargoCuenta = 9
    gAsientoProcesoCreditoSinAsiento = 10
    
    Set RegTran = New ADODB.Recordset
    Set RegTmp = New ADODB.Recordset
    Set RegOpeCta = New ADODB.Recordset
 
    Dim ldFechaAsiento As Date
    
    
    If pAsiDia Then
        ldFechaAsiento = pdfecsis
    Else
        ldFechaAsiento = CDate(dFecha)
    End If
    
    ReDim Arreglo(1, 3)
    'Para la validacion Dia y hora de Generación
    dHoraGrab = Format(pdfecsis & " " & Time, oFunV.gsFormatoFechaHoraView)
    'Carga fecha del Asiento y hora de grabación
    gdHoraGrab = Format(dFecha & " " & Time, oFunV.gsFormatoFechaHora)
    
    Dim vHH As Long, vMM As Long, vSS As Long, vSeg As Long
    Dim vTime As Variant
    vTime = Time
    
    Dim vCta28Debe As Currency, vCta28Haber As Currency
    Dim vCta19Debe As Currency, vCta19Haber As Currency
    Dim vCtaDolar As Currency
    '***********************************************************
    'PARTE 1
    Dim vMonAnt As String
    Dim vTipPer As String
    Dim pband28 As Boolean
    '***********************************************************
    'PARTE 2
    Dim vFondo As String
    Dim vPlazo As String
    Dim vTipoC As String
    Dim pBandFonc As Boolean
    Dim vCJ As String, vCR As String
    Dim vRE As String
    Dim vCE As String
    Dim vBanco As String
    Dim vTipoPD As String
    Dim vSC As String
    Dim vSB As String
    Dim vCD As String
    Dim lsCodAge As String
    Dim vRFA As String
    Dim vAG As String
    '***********************************************************
    
    
    vLenNomb = 70
    vespacio = vLenNomb + 54
    vRTFImp = ""
     
    'Insert registro en tabla para verificación del asiento
    
     InsertaAsientoValida dFecha, pscoduser, dHoraGrab
    
    'Borra el asiento generado de este dia.
    'sSql = "DELETE AsientoDN Where convert(char(10),dfecha,112) = '" & Format(dfecha, "yyyymmdd") & "'"
    
    Dim rsVer As ADODB.Recordset
    Dim lnRegAsi As Long
    
    lnRegAsi = -1
    'Me.Caption = "Eliminando registros de asiento"
    
    Do While lnRegAsi <> 0
     
        EliminarAsiento dFecha
        
        Set rsVer = New ADODB.Recordset
        Set rsVer = ObtenerNroAsientos(dFecha)
        
        If Not rsVer Is Nothing Then
            lnRegAsi = rsVer!nRegAsiento
            rsVer.Close
            Set rsVer = Nothing
        End If
    Loop
    
    'Me.Caption = "Asientos Eliminados correctamente"
    
    'Consulta para obtener los precios del Oro
    
    If Not GetTipCambio(CDate(dFecha)) Then
        psmensaje = " No se encuentran los Tipos de Cambio "
        Exit Function
    Else
        With RegTmp
            TCV = gnTipCambioV
            TCC = gnTipCambioC
            TCF = gnTipCambio
            If TCF = 0 Then
                psmensaje = " No se encuentra el tipo de cambio Fijo Diario"
                RegTmp.Close
                Set RegTmp = Nothing
                Exit Function
            End If
        End With
    End If
    
    'Se carga TranDiaria para generar los asientos de acuerdo ha condiciones ingresadas
    'Para el Asiento
         
    gAsientoProcesoNoAsiento = 3
 
    If pnMonSol = 0 And pnMonDol = 0 Then
        'dbCmact.RollbackTrans ' Corta TRANSACCION
        psmensaje = " No se ha definido la Moneda "
        Exit Function
    End If
    
    Dim nPaso As Long
    vCta28Debe = 0: vCta28Haber = 0
    vCta19Debe = 0: vCta19Haber = 0
    vCont = 0: vCtaCnt = 0
    
    
    'Asiento para Capitalización de Ahorros y CTS
    For nPaso = 1 To 3
        'If pAsiDia Then
        Set RegOpeCta = New ADODB.Recordset
        Set RegOpeCta = ActualizaCapitalizacionAhorrosCTS(nPaso, gAsientoProcesoCapAho, gAsientoProcesoCapCTS, gsCodPersCMACT, ldFechaAsiento)
    
        If RegOpeCta Is Nothing Then
            lsmenError = lsmenError & " - Tipo operación Ahorros no reconocida " & Chr$(10)
            vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * Tipo operación Ahorros no reconocida."
        Else
            With RegOpeCta
                Do While Not .EOF
                    If !cOpeCod = "101301" Then
                        vMoneda = vMoneda
                    End If
                    
                    'Me.Caption = "Capitalizaciones de Ahorros : Registro " & .Bookmark & " de " & .RecordCount
                    
                    'Variables para cambio en el CodContable
                    vMoneda = Trim(!cmoneda)
                    vCodConta = Trim(!cCtaContCod)
                    vCodConta = Replace(vCodConta, "M", vMoneda, , , vbTextCompare)
                    '**** LAYG - 07/01/05 ica
                    lsCodAge = VarAG(!cCtaCod, !cCtaContCod)
                    vAgencia = lsCodAge
                    vCodConta = Replace(vCodConta, "AG", lsCodAge, , , vbTextCompare)
                    '***************
                    If nPaso = 2 Then
                        vCodConta = Replace(vCodConta, "CJ", Trim(!cSubCtaContCod), , , vbTextCompare)
                        vCodConta = Replace(vCodConta, "CJ", Trim(!cSubCtaContCod), , , vbTextCompare)
                        vCodConta = Replace(vCodConta, "CJ", Trim(!cSubCtaContCod), , , vbTextCompare)
                    ElseIf nPaso = 3 Then
                        vCodConta = Replace(vCodConta, "TC", !cTipoCliente, , , vbTextCompare)
                    End If
                    'Se valida la existencia de la Cta. para que se pueda utilizar
                    If Not ExisCtaCnt(vCodConta, True) And Len(Trim(vCodConta)) > 0 Then
                        lsmenError = lsmenError & " - No existe la Cta. Contable y/o no es Cta de Asiento: " & vCodConta & Chr$(10)
                        vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & vCodConta & " cta. cnt. no existe. Operacion: " & !cOpeCod & ")."
                    End If
                        
                    If !cOpeCtaDH = "D" Then
                        InsertaAsientoDN gdHoraGrab, vCodConta, !Monto, 0, "0", vAgencia, !cOpeCod, !cCtaCod

                        If Mid(vCodConta, 3, 1) = Moneda.gMonedaExtranjera Then
                            If Left(vCodConta, 1) = "4" Then
                                InsertaAsientoDN gdHoraGrab, vCodConta, (!Monto * TCV), 0, "3", vAgencia, !cOpeCod, !cCtaCod
                           
                                'Ctas 282503 y 191503
                                vCta28Debe = vCta28Debe + (!Monto * TCF)
                                vCta19Haber = vCta19Haber + (!Monto * TCV)
                            ElseIf Left(vCodConta, 1) = "5" Then
                                InsertaAsientoDN gdHoraGrab, vCodConta, (!Monto * TCC), 0, "3", vAgencia, !cOpeCod, !cCtaCod
    
                                'Ctas 282503 y 191503
                                vCta28Debe = vCta28Debe + (!Monto * TCF)
                                vCta19Haber = vCta19Haber + (!Monto * TCC)
                            Else
                                 InsertaAsientoDN gdHoraGrab, vCodConta, (!Monto * TCF), 0, "3", vAgencia, !cOpeCod, !cCtaCod
                                
                            End If
                        End If
                    ElseIf !cOpeCtaDH = "H" Then
                            InsertaAsientoDN gdHoraGrab, vCodConta, 0, !Monto, "0", vAgencia, !cOpeCod, !cCtaCod
                            
                        If Mid(vCodConta, 3, 1) = Moneda.gMonedaExtranjera Then
                            If Left(vCodConta, 1) = "4" Then
                               
                                InsertaAsientoDN gdHoraGrab, vCodConta, 0, (!Monto * TCV), "3", vAgencia, !cOpeCod, !cCtaCod
                                
                                'Ctas 282503 y 191503
                                vCta28Haber = vCta28Haber + (!Monto * TCF)
                                vCta19Debe = vCta19Debe + (!Monto * TCV)
                            ElseIf Left(vCodConta, 1) = "5" Then
                               
                                InsertaAsientoDN gdHoraGrab, vCodConta, 0, (!Monto * TCC), "3", vAgencia, !cOpeCod, !cCtaCod
                               
                                'Ctas 282503 y 191503
                                vCta28Haber = vCta28Haber + (!Monto * TCF)
                                vCta19Debe = vCta19Debe + (!Monto * TCC)
                            Else
                               
                                InsertaAsientoDN gdHoraGrab, vCodConta, 0, (!Monto * TCF), "3", vAgencia, !cOpeCod, !cCtaCod
                               
                            End If
                        End If
                    Else
                        lsmenError = lsmenError & " - Operación no reconocida al Insertar " & vbCr & _
                        " operación nro.: " & vCodConta & Chr$(10)
                        vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & !cOpeCtaDH & " debe/haber no reconocido (cta.: " & !cOpeCod & ")."
                    End If
                   
                    .MoveNext
                    DoEvents
                Loop
            End With
            RegOpeCta.Close
            Set RegOpeCta = Nothing
       
        End If
    Next
    
    Dim lsRfa As String
    '**** Para la Generación del Asiento
    Set RegTmp = New ADODB.Recordset
    
    Set RegTran = CargaTranDiaria(gAsientoProcesoNoAsientoCab3, gAsientoProcesoNoAsientoCab4, gAsientoProcesoNoAsiento, gAsientoProcesoCreditoSinAsiento, gAsientoProcesoSiAsiento, ldFechaAsiento, gAsientoProcesoSiColocCargoCuenta)
    
    If RegTran Is Nothing Then
    Else
        Do While Not RegTran.EOF
            
            lsRfa = IIf(IsNull(RegTran!cRFA), "", RegTran!cRFA)
            
            'If RegTran!nMovNro = 19065986 Then Stop
            'If RegTran!nmovnro = 19065883 Then Stop
            'If RegTran!nMovNro = 19066085 Then Stop
            
            If RegTran!cOpeCod = "120100" And RegTran!Concepto = "2200" Then
            '    Stop
            End If
                     
            'Variables para cambio en el CodContable
            gdHoraGrab = Format$(RegTran!dFecTran, oFunV.gsFormatoFechaHora)
            ssql = ""
            If Left(RegTran!cOpeCod, 4) = Left(gColPOpeCanceNorEnOtCjEFE, 4) Or _
                Left(RegTran!cOpeCod, 4) = Left(gServGiroCancEfec, 4) Then
                vAgencia = Left(RegTran!cCtaCod, 2)
            Else
                If Len(RegTran!cCodAge) > 2 Then
                    vAgencia = Mid(RegTran!cMovNro, 18, 2)
                Else
                    vAgencia = Right(Trim(RegTran!cCodAge), 2)
                End If
            End If
            
            If Not IsNull(RegTran!cCtaCod) Then
                'Verifica Moneda de Cuentas Antiguas o Cuentas Nuevas
                If Len(Trim(RegTran!cCtaCod)) = 8 Then
                    vMonAnt = Mid(Trim(RegTran!cCtaCod), 2, 1)
                    If vMonAnt = "0" Then
                        vMoneda = Moneda.gMonedaNacional
                    ElseIf vMonAnt = Moneda.gMonedaNacional Then
                        vMoneda = Moneda.gMonedaExtranjera
                    Else
                        lsmenError = lsmenError & " - Moneda Antigua no reconocida " & RegTran!cCtaCod & Chr$(10)
                        vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " moneda antigua no reconocida."
                    End If
                Else
                    vMoneda = Mid(Trim(RegTran!cCtaCod), 9, 1)
                    If Not (vMoneda = Moneda.gMonedaNacional Or vMoneda = Moneda.gMonedaExtranjera) Then
                        lsmenError = lsmenError & " - Código Errado, Moneda no definida " & RegTran!cCtaCod & Chr$(10)
                        vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " moneda no reconocida."
                    End If
                End If
                'En ahorros se busca el tipo de persona en ctas de Ahorro
                
                If (Left(Trim(RegTran!cOpeCod), 1) = "2" Or Trim(RegTran!cOpeCod) = "990101" Or Trim(RegTran!cOpeCod) = "990301") _
                    And ExisTipPer(RegTran!cOpeCod, RegTran!Concepto, True) Then 'RERR: agregue las operaciones de ITF Cargo cuenta
                    vTipPer = ClienteTipoPers(RegTran!cCtaCod)
                    If Len(Trim(vTipPer)) = 0 Then
                        lsmenError = lsmenError & " - No existe Titular, Operac. Ahorros " & RegTran!cCtaCod & Chr$(10)
                        vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " no existe Titular (Operac. Ahorros)."
                    End If
                    ssql = " AND nPersoneria = '" & vTipPer & "'"
                'En Crédito se busca Tipo de Persona en algunas Cuentas de Créditos
                'MARCA 01
                ElseIf Left(Trim(RegTran!cOpeCod), 4) = "0111" Or Left(Trim(RegTran!cOpeCod), 4) = "0118" Or _
                    Left(Trim(RegTran!cOpeCod), 4) = "0104" Or Left(Trim(RegTran!cOpeCod), 4) = "0132" Or _
                    Left(Trim(RegTran!cOpeCod), 4) = "0125" Or Left(Trim(RegTran!cOpeCod), 4) = "0139" Or _
                    Left(Trim(RegTran!cOpeCod), 4) = "0181" Or Left(Trim(RegTran!cOpeCod), 6) = "100202" Or _
                    Trim(RegTran!cOpeCod) = "100102" Or Trim(RegTran!cOpeCod) = "100103" Or Trim(RegTran!cOpeCod) = "100302" Or _
                    Trim(RegTran!cOpeCod) = "100104" Or Trim(RegTran!cOpeCod) = "100104" Or _
                    Trim(RegTran!cOpeCod) = "100105" Or Trim(RegTran!cOpeCod) = "100106" Or _
                    Trim(RegTran!cOpeCod) = "100107" Or Trim(RegTran!cOpeCod) = "100108" Or _
                    Trim(RegTran!cOpeCod) = "100109" Or Left(Trim(RegTran!cOpeCod), 4) = "0188" Then
                        
                        If Left(Trim(RegTran!cOpeCod), 6) = "010440" Or Left(Trim(RegTran!cOpeCod), 4) = "0188" Or _
                            Left(Trim(RegTran!cOpeCod), 6) = "010105" Then
                            
                            vTipPer = ClienteTipoPersCol(RegTran!nMovNro)
                            
                        Else
                            If Left(RegTran!cOpeCod, 5) = "01043" Or Left(RegTran!cOpeCod, 5) = "01113" Or _
                            Left(RegTran!cOpeCod, 5) = "01183" Or Left(RegTran!cOpeCod, 5) = "01253" Or _
                            Left(RegTran!cOpeCod, 5) = "01393" Then
                                vTipPer = 0
                            Else
                                vTipPer = ClienteTipoPersCol(RegTran!nMovNro)
                            End If
                        End If
                        If Len(Trim(vTipPer)) = 0 Then
                            lsmenError = lsmenError & " - No existe Titular, Operac. Crédito " & RegTran!cCtaCod & Chr$(10)
                            vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " no existe Titular (Operac. Créditos)."
                        End If
                        ssql = " AND nPersoneria = " & vTipPer & " "
                ElseIf Left(Trim(RegTran!cOpeCod), 2) = Left(gColRecOpePasoARecup, 2) Then ' Arturo - Judicial
                    If RegTran!nSaldCnt = 1 Or RegTran!nSaldCnt = 2 Or RegTran!nSaldCnt = 3 Or RegTran!nSaldCnt = 4 Then
                        ssql = " AND nPersoneria = '" & RegTran!nSaldCnt & "'"
                    End If
                End If
            Else
                If Not (Left(Trim(RegTran!cOpeCod), 2) = "23" Or Left(Trim(RegTran!cOpeCod), 2) = "24" Or Left(Trim(RegTran!cOpeCod), 2) = "25") Then
                    lsmenError = lsmenError & " - No es un Código de Operac. Especial " & RegTran!cOpeCod & Chr$(10)
                    vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cOpeCod & " no es un código de Operac. Especial (Cta.: " & RegTran!cCtaCod & ")."
                End If
            End If
            'Para Asiento de Crédito, Foncodes pago en otra agencia
            pband28 = True
            
            'If RegTran!cOpeCod = "100505" Then Stop
                    
            'Carga OPECUENTA para definir a que cuentas se ha de grabar
            Set RegOpeCta = New ADODB.Recordset
            Set RegOpeCta = CargaOpeCuenta(RegTran!cOpeCod, RegTran!Concepto, ssql)
            
            If RegOpeCta Is Nothing Then
                If (Left(RegTran!cOpeCod, 3) <> "107" And RegTran!cOpeCod <> "120100") Then
                    lsmenError = lsmenError & " - Código de Operación No Reconocido " & RegTran!cOpeCod & "- Concepto " & RegTran!Concepto & Chr$(10)
                End If
                If RegTran!cOpeCod <> "120100" Then 'OBVIA LOS DEMAS CONCEPTOS DEL REGISTRO DE CREDITOS PIGNORATICIO
                    vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cOpeCod & " operación no reconocida (Cta.: " & RegTran!cCtaCod & "- Concepto " & RegTran!Concepto & ")."
                End If
            Else
                Do While Not RegOpeCta.EOF
                    'Mensaje de Cuenta
                    vCtaCnt = vCtaCnt + 1
                    If Len(RegTran!cCodAge) > 2 Then
                        vAgencia = Mid(RegTran!cMovNro, 18, 2)
                    Else
                        vAgencia = Right(Trim(RegTran!cCodAge), 2)
                    End If
                    
                    vSeg = DateDiff("s", vTime, Time)
                    vSS = vSeg - (Int(vSeg / 60) * 60)
                    vMM = Int(vSeg / 60)
                    vHH = Int(vSeg / 3600)
                    
                    'Me.Caption = " CodCta.: " & RegTran!cCtaCod & " - Regis.: " & vCont + 1 & " - CtaCnt.: " & vCtaCnt & " - Time -> " & FillNum(Str(vHH), 6, "0") & ":" & FillNum(Str(vMM), 6, "0") & ":" & FillNum(Str(vSS), 6, "0")
                    
                    vCodConta = GetPlantillaPuente(RegOpeCta!cCtaContCod, Mid(RegTran!cCtaCod, 6, 3), RegTran!Concepto, RegTran!cOpeCod)
                   
                    If vCodConta = "" Then
                        vCodConta = Trim(RegOpeCta!cCtaContCod)
                    End If
                    If Not IsNull(RegTran!cCtaCod) Then
                        'Variable de cambios
                        vFondo = "": vPlazo = "": vTipoC = "": vRFA = ""
                        ''' validacion de creditos RFA (RFC-DIF)
                        '**14M4TC2501RFCD0PSCFFAG
                        If (lsRfa = "RFC" Or lsRfa = "DIF") And Left(vCodConta, 2) = "14" Then
                            'Stop
                            '' CAMBIAMOS PLANTILLAS PARA LOS CREDITOS rfa SI ES CUENTA 14
                            'If RegTran!nCredEstado = 2031 Then
                            If Left(vCodConta, 4) = "14M5" Then
                                vCodConta = Left(vCodConta, 6) + "1902RFCD0PSCFFAG"
                            Else
                                vCodConta = Left(vCodConta, 6) + "2501RF060PSCFFAG"
                            End If
                            'tipo de credito RFA (NEMMONICO RF)
                            Select Case lsRfa
                                Case "RFC"
                                    vRFA = "01"
                                Case "DIF"
                                    vRFA = "02"
                                Case Else
                                    vRFA = ""
                                    vCodConta = "15"
                            End Select
                            vCodConta = Replace(vCodConta, "RF", vRFA, , , vbTextCompare)
                        End If
                        'arreglo de creditos de consumo vencidos refinanciados - EJRS
                        If (Left(vCodConta, 4) = "14M5" Or Left(vCodConta, 4) = "14M6") And Left(RegTran!cOpeCod, 4) = "1007" And Mid(RegTran!cCtaCod, 6, 1) = "3" Then
                            'ELIMINAMOS EL PAR 29 PUESTO QUE EN CONSUMO NO SE CONSIDERA
                            vCodConta = Replace(vCodConta, "29", "")
                        End If
                        
                        If Left(vCodConta, 4) = "51M4" And Mid(RegTran!cCtaCod, 6, 3) = "320" Then
                            vCodConta = Replace(vCodConta, "FF", "")
                        End If
                        
                        'Tipo de credito
                        If InStr(vCodConta, "TC") > 0 Then
                            If Mid(RegTran!cCtaCod, 6, 3) = Producto.gCapCTS Then
                                vTipoC = ClienteTipoCTS(RegTran!cCtaCod)
                                If vTipoC = "01" Or vTipoC = "02" Then
                                    vCodConta = Replace(vCodConta, "TC", vTipoC, , , vbTextCompare)
                                Else
                                    lsmenError = lsmenError & " - Tipo de Cliente - CTS Errado " & RegTran!cCtaCod & Chr$(10)
                                    vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " tipo de cliente - CTS errado."
                                End If
                            Else ' Para TC Tipo de Creditos
                                vTipoC = Mid(RegTran!cCtaCod, 6, 1)
                                vCodConta = Replace(vCodConta, "TC", Trim("0" & vTipoC), , , vbTextCompare)
                            End If
                        End If
                        
                        If InStr(vCodConta, "SC") > 0 Then
                            ' Cambiado LAYG 03/03/03 Usar la tabla AsientoProdEquiv
                            If Mid(RegTran!cCtaCod, 6, 3) = "320" Then
                                'enrique
                                vSC = VarAsientoPrestamosAdm(RegTran!cCtaCod)
                            Else
                                vSC = VarAsientoProdEquiv(Mid(RegTran!cCtaCod, 6, 3), "SC")
                            End If
                            
                            If vSC = "" Then
                                lsmenError = lsmenError & " - Tipo de SC Errado " & RegTran!cCtaCod & Chr$(10)
                                vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " tipo de SC errado."
                            Else
                                vCodConta = Replace(vCodConta, "SC", vSC, , , vbTextCompare)
                            End If
                        End If
                        
                        If InStr(vCodConta, "CD") > 0 Then
                        '   If Mid(RegTran!cCtaCod, 7, 2) = "20" Then
                        '      vCodConta = Replace(vCodConta, "CD", "20", , , vbTextCompare)
                        '   Else
                        '      vCodConta = Replace(vCodConta, "CD", "06", , , vbTextCompare)
                        '   End If
                            ' Cambiado LAYG 03/03/03 Usar la tabla AsientoProdEquiv
                           
                            vCD = VarAsientoProdEquiv(Mid(RegTran!cCtaCod, 6, 3), "CD")
                            
                            If vCD = "" Then
                                lsmenError = lsmenError & " -  Tipo de CD Errado " & RegTran!cCtaCod & Chr$(10)
                                vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " tipo de SC errado."
                            Else
                                vCodConta = Replace(vCodConta, "CD", vCD, , , vbTextCompare)
                            End If
                        End If
                        
                        'Cambios para FF - Fondos
                        If InStr(vCodConta, "FF") > 0 Then
                            
                            vFondo = VarFF(RegTran!cCtaCod, IIf(IsNull(RegTran!LineaC), "", RegTran!LineaC), "")
                            
                            vCodConta = Replace(vCodConta, "FF", vFondo, 1, 1, vbTextCompare)
                        End If
                        'MARCA 02 ¿Que es fondo?
                        'Valida sólo para Crédito FF:02 - LC:01 y Cta = "28...."
                        If Left(Trim(RegTran!cOpeCod), 2) = "10" And vFondo = "02" _
                            And Left(vCodConta, 2) = "29" And Left(Trim(RegTran!cOpeCod), 3) <> "018" Then
                            vCodConta = ""
                            pband28 = False
                        End If
                        'Valida para Foncodes - Crédito
                        'MARCA 04
                        pBandFonc = False
                        If Left(Trim(RegTran!cOpeCod), 2) = "10" And vFondo = "02" And _
                            (Left(vCodConta, 2) = "14" Or Left(vCodConta, 2) = "51") Then
                            pBandFonc = True
                        End If
                        'Cambios para AO
                        'MARCA 05
                        If InStr(vCodConta, "AO") > 0 Then
                            If InStr(RegTran!cNumDoc, "@") = 1 Then
                                vCodConta = Replace(vCodConta, "AO", Mid(RegTran!cNumDoc, 5, 2), , , vbTextCompare)
                            Else
                                lsmenError = lsmenError & " -  No se encontro AO - Créditos " & RegTran!cCtaCod & Chr$(10)
                                vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " no se encontro AO - Créditos."
                            End If
                        End If
                    End If
                    'Codificacion de los bancos para el canje de Ordenes de Pago
                    If InStr(vCodConta, "BC") > 0 Then
                        vBanco = VarBC(RegTran!nMovNro, vCodConta)
                        If vBanco = "" Then
                            lsmenError = lsmenError & " - Número de Banco no Identificado (Después de función) " & RegTran!cCtaCod & Chr$(10)
                            vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " número de banco no identificado (Después de función) ."
                        Else
                            vCodConta = Replace(vCodConta, "BC", vBanco)
                        End If
                    End If
'                    If InStr(vCodConta, "SB") > 0 Then
'                        If RegTran!nNroCuota <> 0 And Trim(RegTran!nNroCuota) <> "" Then
'                            vSB = VarBC(RegTran!nNroCuota, , True)
'                            vCodConta = Replace(vCodConta, "SB", vSB)
'                        Else
'                            MsgBox " Número de SB no Identificado (Antes de función) " & RegTran!cCtaCod, vbInformation, " Aviso "
'                            vNoCtaCnt = vNoCtaCnt & CHR$(10) & "     * " & RegTran!cCtaCod & " número de SB no identificado (Antes de función) ."
'                        End If
'                    End If
                    'Valida la CJ - Ahorros - Credito - Pignoraticio
                    'MARCA 06
                    If InStr(vCodConta, "CJ") > 0 Then
                        '************ OJO HAY QUE VERIFICAR ***************************
                     
                        Set oPers = New COMDPersona.DCOMInstFinac
                            vCJ = oPers.VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCmac)
                        Set oPers = Nothing
                        
                        Set oPers = New COMDPersona.DCOMInstFinac
                        If vCJ = "" Then
                           vCJ = oPers.VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCooperativa)
                        End If
                        If vCJ = "" Then
                           vCJ = oPers.VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFEDPYME)
                        End If
                        If vCJ = "" Then
                           vCJ = oPers.VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCrac)
                        End If
                        Set oPers = Nothing
                        
                        If vCJ = "" Then
                            ' parche para operaciones de recepcion CMAC
                            If (Left(RegTran!cOpeCod, 4) >= "2601" And Left(RegTran!cOpeCod, 4) <= "2603") Or _
                                RegTran!cOpeCod = "100205" Or Left(RegTran!cOpeCod, 4) = "9903" Or RegTran!cOpeCod = "100405" Or _
                                RegTran!cOpeCod = "100305" Or RegTran!cOpeCod = "136301" Or RegTran!cOpeCod = "136201" Or _
                                Left(RegTran!cOpeCod, 3) = "126" Or RegTran!cOpeCod = "100505" Then
                                
                                'Los 126 Agregado por JHVP
                                Set oPers = New COMDPersona.DCOMInstFinac
                                    vCJ = oPers.VarInstFinanMov(RegTran!nMovNro, gTpoIFCmac)
                                Set oPers = Nothing
                                
                                If vCJ = "" Then
                                    lsmenError = lsmenError & " - CJ no reconocido (Ahorros) " & RegTran!cCtaCod & Chr$(10)
                                    vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " CJ no reconocido (Ahorros)."
                                Else
                                    vCodConta = Replace(vCodConta, "CJ", vCJ, 1, 1, vbTextCompare)
                                End If
                            Else
                                If Left(RegTran!cOpeCod, 1) = "2" Or Left(RegTran!cOpeCod, 2) = "31" Then
                                    vCodConta = "19M901CJAG"
                                ElseIf Left(RegTran!cOpeCod, 2) = "01" Or Left(RegTran!cOpeCod, 2) = "33" Or Left(RegTran!cOpeCod, 2) = "34" Or Left(RegTran!cOpeCod, 2) = "35" Then
                                    vCodConta = "29M80701PDAG"
                                ElseIf Left(RegTran!cOpeCod, 2) = "03" Then
                                    vCodConta = "29M8070104AG"
                                Else
                                    lsmenError = lsmenError & " - CJ no reconocido (Ahorros) " & RegTran!cCtaCod & Chr$(10)
                                    vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " CJ no reconocido (Ahorros)."
                                End If
                            End If
                        Else
                            vCodConta = Replace(vCodConta, "CJ", vCJ, 1, 1, vbTextCompare)
                        End If
                    End If
                    'Valida la CR - Otras Instituciones Financieras - Cajas Rurales (Raul)
                    If InStr(vCodConta, "CR") > 0 Then
                        Set oPers = New COMDPersona.DCOMPersona
                            vCR = oPers.VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCrac)
                        Set oPers = Nothing
                        If vCR = "" Then
                            lsmenError = lsmenError & " - CR no reconocido (Ahorros) " & RegTran!cCtaCod & Chr$(10)
                            vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " CR no reconocido (Ahorros)."
                        Else
                            vCodConta = Replace(vCodConta, "CR", vCR, 1, 1, vbTextCompare)
                        End If
                    End If
                    If InStr(vCodConta, "CO") > 0 Then
                        Set oPers = New COMDPersona.DCOMPersona
                            vCR = oPers.VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFCooperativa)
                        Set oPers = Nothing
                        If vCR = "" Then
                            lsmenError = lsmenError & " - CO no reconocido (Ahorros) " & RegTran!cCtaCod & Chr$(10)
                            vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " CR no reconocido (Ahorros)."
                        Else
                            vCodConta = Replace(vCodConta, "CO", vCR, 1, 1, vbTextCompare)
                        End If
                    End If
                    If InStr(vCodConta, "ED") > 0 Then
                        Set oPers = New COMDPersona.DCOMPersona
                            vCR = oPers.VarInstitucionFinanciera(RegTran!cCtaCod, gTpoIFEDPYME)
                        Set oPers = Nothing
                        If vCR = "" Then
                            lsmenError = lsmenError & " -  ED no reconocido (Ahorros) " & RegTran!cCtaCod & Chr$(10)
                            vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " CR no reconocido (Ahorros)."
                        Else
                            vCodConta = Replace(vCodConta, "ED", vCR, 1, 1, vbTextCompare)
                        End If
                    End If
                    
                    'Créditos
                    If InStr(vCodConta, "PD") > 0 Then
                        vTipoPD = VarPD(RegTran!cCtaCod, IIf(IsNull(RegTran!LineaC), "", RegTran!LineaC), "")
                        If vTipoPD = "01" Or vTipoPD = "02" Or vTipoPD = "03" Or vTipoPD = "05" Then
                            vCodConta = Replace(vCodConta, "PD", vTipoPD, , , vbTextCompare)
                        Else
                            lsmenError = lsmenError & "  - Producto Pendiente Errado " & RegTran!cCtaCod & Chr$(10)
                            vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " producto pendiente errado."
                        End If
                    End If
    
                    'Cambios para P - Plazo
                    If InStr(vCodConta, "P") > 0 Then
                        vPlazo = Mid(Trim(RegTran!LineaC), 6, 1)
                        If vPlazo = "1" Or vPlazo = "2" Then
                            vCodConta = Replace(vCodConta, "P", vPlazo, 1, 1, vbTextCompare)
                        Else
                            lsmenError = lsmenError & " - Plazo de Crédito Errada " & RegTran!cCtaCod & Chr$(10)
                            vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " plazo de crédito errado."
                        End If
                    End If
                    
                    'Valida la agencia remota - AG - AL - AR
                    If InStr(vCodConta, "AR") > 0 Then
                        vCodConta = Replace(vCodConta, "AR", vAgencia, 1, 1, vbTextCompare)
                    End If
                    '------ LAYG 07/01/2005 ica
                    If InStr(vCodConta, "AG") > 0 Then
                        If Left(vCodConta, 2) = "11" Then
                            'vCodConta = Replace(vCodConta, "AG", Mid(RegTran!cMovNro, 15, 5), 1, 1, vbTextCompare)
                            vAG = VarAG(Mid(RegTran!cMovNro, 15, 5), vCodConta)
                        Else
                            If Left(RegTran!cOpeCod, 4) = Left(gPigOpeDespContNuevo, 4) Then
                                'vCodConta = Replace(vCodConta, "AG", vAgencia, 1, 1, vbTextCompare)
                                vAG = VarAG(pscodcmac & vAgencia, vCodConta)
                            Else
                                'vCodConta = Replace(vCodConta, "AG", Mid(RegTran!cCtaCod, 4, 2), 1, 1, vbTextCompare)
                                vAG = VarAG(RegTran!cCtaCod, vCodConta)
                            End If
                        End If
                        'CAMBIAMOS AGENCIA SI OPERACION ES DESEMBOLSO CON ABONO A CUENTA DE OTRA AGENCIA
                        'NUEVO PARCHE EJRS  26 DE MAYO 2005
                        If RegTran!cOpeCod = "100104" And Left(vCodConta, 2) = "21" Then
                           vAG = GetAgenciaCtaDesmAbonoCta(RegTran!nMovNro, RegTran!cOpeCod)
                           vAgencia = vAG
                        End If
                        If vAG = "" Then
                            lsmenError = lsmenError & " - AG no reconocida " & RegTran!cCtaCod & Chr$(10)
                            vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegTran!cCtaCod & " AG no reconocida."
                        Else
                            vCodConta = Replace(vCodConta, "AG", vAG, 1, 1, vbTextCompare)
                        End If
                        
                    End If
                    '----------
                    'Cambios en Moneda y Agencia
                    vCodConta = Replace(vCodConta, "M", vMoneda, 1, 1, vbTextCompare)
                    'Asiento para ctas. de Foncodes en la Cta de Ahorro : 012321006254 - RAUL
                    'MARCA 08
                    'If RegTran!cCtaCod = gsCtaCodFoncodes And Left(Trim(vCodConta), 10) = "2112010102" Then
                        'vCodConta = "211201010290"
                    'End If
                    'Para unificar varias cuentas en una - ARTURO
                    
                    vParche = AsientoParche(vCodConta, True)
                   
                    
                    If Len(vParche) > 0 Then
                        vCodConta = vParche
                    End If
                    'Se valida la existencia de la Cta. para que se pueda utilizar
                    If (Not ExisCtaCnt(vCodConta, True) And Len(Trim(vCodConta)) > 0) Or Len(Trim(vCodConta)) = 0 Then
                        lsmenError = lsmenError & " - No existe la Cta. Contable y/o no es Cta de Asiento: " & vCodConta & Chr$(10)
                        vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & vCodConta & " cta. cnt. no existe (cta.: " & RegTran!cCtaCod & ", ope.: " & RegTran!cOpeCod & " - RFA = " & RegTran!cRFA & " - Concepto " & RegTran!Concepto & ")."
                    End If
                    
                    'If Trim(vCodConta) = "19180207" Then Stop
                    
                    '****************************************************************************
                    'Verifica si el asiento ya fue creado
                    'Si existe lo actualiza; caso contrario lo agrega
                    If Len(Trim(vCodConta)) > 0 Then
                        If RegOpeCta!cOpeCtaDH = "D" Then
                           
                            InsertaAsientoDN gdHoraGrab, vCodConta, RegTran!sumaMonto, 0, IIf(RegTran!Concepto = 8 Or RegTran!Concepto = 9, "3", "0"), vAgencia, RegTran!cOpeCod, RegTran!cCtaCod, RegTran!nMovNro
                            
                                    
                            If Mid(vCodConta, 3, 1) = Moneda.gMonedaExtranjera And Not (RegTran!Concepto = 8 Or RegTran!Concepto = 9) Then
                                If Left(vCodConta, 1) = "4" Then
                                   
                                    InsertaAsientoDN gdHoraGrab, vCodConta, (RegTran!sumaMonto * TCV), 0, " 3", vAgencia, RegTran!cOpeCod, RegTran!cCtaCod, RegTran!nMovNro
                                                    
                                    'Ctas 282503 y 191503
                                    vCta28Debe = vCta28Debe + (RegTran!sumaMonto * TCF)
                                    vCta19Haber = vCta19Haber + (RegTran!sumaMonto * TCV)
                                ElseIf Left(vCodConta, 1) = "5" Then
                                   
                                    InsertaAsientoDN gdHoraGrab, vCodConta, (RegTran!sumaMonto * TCC), 0, "3", vAgencia, RegTran!cOpeCod, RegTran!cCtaCod, RegTran!nMovNro
                       
                                    'Ctas 282503 y 191503
                                    vCta28Debe = vCta28Debe + (RegTran!sumaMonto * TCF)
                                    vCta19Haber = vCta19Haber + (RegTran!sumaMonto * TCC)
                                Else
                                    
                                    InsertaAsientoDN gdHoraGrab, vCodConta, (RegTran!sumaMonto * TCF), 0, "3", vAgencia, RegTran!cOpeCod, RegTran!cCtaCod, RegTran!nMovNro
                                    
                                End If
                            End If
                            'Solo Inserta cuando es Foncodes y la bandera está en True cta-14  y cta sea diferente a Cambio de Cartera
                            
                            'MARCA 09
'                            If pBandFonc And Left(RegTran!cOpeCod, 4) <> "0160" And _
'                            (Left(vCodConta, 2) = "14" Or Left(vCodConta, 2) = "51") Then
'                                vCodConta = "26M60901"
'                                vCodConta = Replace(vCodConta, "M", vMoneda, 1, 1, vbTextCompare)
'                                    sSql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                        " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & RegTran!sumaMonto & ",'0','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpeCod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar sSql
'                                    If Mid(vCodConta, 3, 1) = Moneda.gMonedaExtranjera Then
'                                        sSql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                            " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & (RegTran!sumaMonto * TCF) & ",'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpeCod & "','" & RegTran!cCtaCod & "')"
'                                        oCon.Ejecutar sSql
'                                    End If
'                                If pband28 Then
'                                    vCodConta = "29M807010501"
'                                    vCodConta = Replace(vCodConta, "M", vMoneda, 1, 1, vbTextCompare)
'                                    vCodConta = Replace(vCodConta, "AG", vAgencia, 1, 1, vbTextCompare)
'                                        sSql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                            " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & RegTran!sumaMonto & ",0,'0','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpeCod & "','" & RegTran!cCtaCod & "')"
'                                        oCon.Ejecutar sSql
'                                        If Mid(vCodConta, 3, 1) = Moneda.gMonedaExtranjera Then
'                                            sSql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                                " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & (RegTran!sumaMonto * TCF) & ",0,'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpeCod & "','" & RegTran!cCtaCod & "')"
'                                            oCon.Ejecutar sSql
'                                        End If
'                                End If
'                            End If
                        ElseIf RegOpeCta!cOpeCtaDH = "H" Then
                           
                            InsertaAsientoDN gdHoraGrab, vCodConta, 0, RegTran!sumaMonto, IIf(RegTran!Concepto = 8 Or RegTran!Concepto = 9, "3", "0"), vAgencia, RegTran!cOpeCod, RegTran!cCtaCod, RegTran!nMovNro
                         
                            If Mid(vCodConta, 3, 1) = Moneda.gMonedaExtranjera And Not (RegTran!Concepto = 8 Or RegTran!Concepto = 9) Then
                                If Left(vCodConta, 1) = "4" Then
 
                                    InsertaAsientoDN gdHoraGrab, vCodConta, 0, (RegTran!sumaMonto * TCV), IIf(RegTran!Concepto = 8 Or RegTran!Concepto = 9, "3", "0"), vAgencia, RegTran!cOpeCod, RegTran!cCtaCod, RegTran!nMovNro
                           
                                    'Ctas 282503 y 191503
                                    vCta28Haber = vCta28Haber + (RegTran!sumaMonto * TCF)
                                    vCta19Debe = vCta19Debe + (RegTran!sumaMonto * TCV)
                                ElseIf Left(vCodConta, 1) = "5" Then
                                 
                                    InsertaAsientoDN gdHoraGrab, vCodConta, 0, (RegTran!sumaMonto * TCC), IIf(RegTran!Concepto = 8 Or RegTran!Concepto = 9, "3", "0"), vAgencia, RegTran!cOpeCod, RegTran!cCtaCod, RegTran!nMovNro
                                    
                                    'Ctas 282503 y 191503
                                    vCta28Haber = vCta28Haber + (RegTran!sumaMonto * TCF)
                                    vCta19Debe = vCta19Debe + (RegTran!sumaMonto * TCC)
                                Else
                                    
                                    InsertaAsientoDN gdHoraGrab, vCodConta, 0, (RegTran!sumaMonto * TCF), IIf(RegTran!Concepto = 8 Or RegTran!Concepto = 9, "3", "0"), vAgencia, RegTran!cOpeCod, RegTran!cCtaCod, RegTran!nMovNro
                                    
                                End If
                            End If
                            'Solo Inserta cuando es Foncodes y la bandera está en True cta-14  y cta sea diferente a Cambio de Cartera
                            'MARCA 10
'                            If pBandFonc And Left(RegTran!cOpeCod, 4) <> "0160" And _
'                            (Left(vCodConta, 2) = "14" Or Left(vCodConta, 2) = "51") Then
'                                vCodConta = "26M60901"
'                                vCodConta = Replace(vCodConta, "M", vMoneda, 1, 1, vbTextCompare)
'                                    sSql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                        " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & RegTran!sumaMonto & ",0,'0','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpeCod & "','" & RegTran!cCtaCod & "')"
'                                    oCon.Ejecutar sSql
'                                    If Mid(vCodConta, 3, 1) = Moneda.gMonedaExtranjera Then
'                                        sSql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                            " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & (RegTran!sumaMonto * TCF) & ",0,'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpeCod & "','" & RegTran!cCtaCod & "')"
'                                        oCon.Ejecutar sSql
'                                    End If
'                                If pband28 Then
'                                    vCodConta = "29M807010501"
'                                    vCodConta = Replace(vCodConta, "M", vMoneda, 1, 1, vbTextCompare)
'                                    vCodConta = Replace(vCodConta, "AG", vAgencia, 1, 1, vbTextCompare)
'                                        sSql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                            " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & RegTran!sumaMonto & ",'0','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpeCod & "','" & RegTran!cCtaCod & "')"
'                                        oCon.Ejecutar sSql
'                                        If Mid(vCodConta, 3, 1) = Moneda.gMonedaExtranjera Then
'                                            sSql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) " & _
'                                                " VALUES('" & gdHoraGrab & "','" & vCodConta & "',0," & (RegTran!sumaMonto * TCF) & ",'3','" & vAgencia & "'," & RegTran!nMovNro & ",'" & RegTran!cOpeCod & "','" & RegTran!cCtaCod & "')"
'                                            oCon.Ejecutar sSql
'                                        End If
'                                End If
'                            End If
                        Else
                            lsmenError = lsmenError & " - Operación no reconocida al Insertar " & vbCr & _
                            " operación nro.: " & vCodConta & Chr$(10)
                            vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegOpeCta!cOpeCtaDH & " debe/haber no reconocido (cta.: " & RegTran!cCtaCod & ")."
                        End If
                    End If
                    RegOpeCta.MoveNext
                Loop
                RegOpeCta.Close
                Set RegOpeCta = Nothing
            End If
            DoEvents
            vCont = vCont + 1
            
            RegTran.MoveNext
        Loop
            
            
'        'Verifica que cuadre DEBE y HABER - Tipo '3'
'        Dim vDife As Currency, vDifeSol As Currency
'        Dim vMonDebe As Currency, vMonHaber As Currency
'        Dim vMonDebeSol As Currency, vMonHaberSol As Currency
'        'Carga suma de tipo soles para hallar la diferencia - RAUL
'        tmpSql = "SELECT round(SUM(ndebe),2) AS MonDebe , round(SUM(nhaber),2) AS MonHaber FROM AsientoDN " & _
'            " WHERE ctipo = '0' AND substring(cCtaCnt,3,1) = '" & Moneda.gMonedaNacional & "' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
'            " Group BY cCtaCnt "
'        vMonDebeSol = 0: vMonHaberSol = 0: vMonDebe = 0: vMonHaber = 0
'        Set RegTmp = oCon.CargaRecordSet(tmpSql)
'        If (RegTmp.BOF Or RegTmp.EOF) Then
'            'No se encuentra ctas en Dolares o Tipo 3
'        Else
'            Do While Not RegTmp.EOF
'                vMonDebeSol = vMonDebeSol + RegTmp!mondebe
'                vMonHaberSol = vMonHaberSol + RegTmp!monhaber
'                RegTmp.MoveNext
'            Loop
'        End If
'        RegTmp.Close
'        Set RegTmp = Nothing
'        'Halla suma de diferencia de Dolares
'        tmpSql = "SELECT round(SUM(ndebe),2) AS MonDebe , round(SUM(nhaber),2) AS MonHaber FROM AsientoDN " & _
'            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & gdHoraGrab & "') = 0 " & _
'            " Group BY cCtaCnt "
'        Set RegTmp = oCon.CargaRecordSet(tmpSql)
'        If (RegTmp.BOF Or RegTmp.EOF) Then
'            'No se encuentra ctas en Dolares o Tipo 3
'        Else
'            Do While Not RegTmp.EOF
'                vMonDebe = vMonDebe + RegTmp!mondebe
'                vMonHaber = vMonHaber + RegTmp!monhaber
'                RegTmp.MoveNext
'            Loop
'        End If
'        RegTmp.Close
'        Set RegTmp = Nothing
'
'        vDifeSol = vMonDebeSol - vMonHaberSol
'
'        vMonDebe = Round(vMonDebe, 2)
'        vMonHaber = Round(vMonHaber, 2)
'        vDife = vMonDebe - vMonHaber
'
'        vDife = vDife + vDifeSol
'        'Inserta a la cuenta de ganancia (51280401-Haber) o perdida (41180401-Debe)
'        'vCtaDolar = (vCta19Debe + vCta28Debe) - (vCta19Haber + vCta28Haber)
'
'        vCtaDolar = -1 * vDife 'CtaDolar + vDife
'
'
'        If vCtaDolar < 0 Then
'            sSQL = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
'                " VALUES('" & gdHoraGrab & "','" & lsCVMEGanacia & "',0," & Abs(vCtaDolar) & ",'3','" & vAgencia & "')"
'            oCon.Ejecutar sSQL
'        ElseIf vCtaDolar > 0 Then
'            sSQL = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
'                " VALUES('" & gdHoraGrab & "','" & lsCVMEPerdida & "'," & Abs(vCtaDolar) & ",0,'3','" & vAgencia & "')"
'            oCon.Ejecutar sSQL
'        End If
        '**************************************************************************************
        'Redondea la conversion de Dolares (Tipo 3)
        'SSQL = "UPDATE AsientoDN SET ndebe = round(ndebe,2) , nhaber = round(nhaber,2) " & _
            " WHERE ctipo = '3' AND datediff(dd, dFecha ,'" & Format(gdHoraGrab, oFunV.gsFormatoFecha) & "') = 0"
        'oCon.Ejecutar SSQL
        '**************************************************************************************
        Dim vCta83D As Currency, vCta83H As Currency, vCta84D As Currency, vCta84H As Currency, vDife8 As Currency
        'CUADRE DE CTAS DE ORDEN (SILVITA)   83 = 84
        'MARCA 11 CUENTAS DE ORDEN
        vCta83D = 0: vCta83H = 0
        Set RegTmp = New ADODB.Recordset
      
        Set RegTmp = CuadreCtas(gdHoraGrab, "83%")
      
        
        If RegTmp Is Nothing Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta83D = vCta83D + RegTmp!nDebe
                vCta83H = vCta83H + RegTmp!nhaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta83D = Abs(vCta83D)
        vCta83H = Abs(vCta83H)
        
        vCta84D = 0: vCta84H = 0
        Set RegTmp = New ADODB.Recordset
        Set RegTmp = CuadreCtas(gdHoraGrab, "84%")
        If RegTmp Is Nothing Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta84D = vCta84D + RegTmp!nDebe
                vCta84H = vCta84H + RegTmp!nhaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta84D = Abs(vCta84D)
        vCta84H = Abs(vCta84H)
        
        'Validación
        'MsgBox (vCta83D + vCta83H) & "  " & (vCta84D + vCta84H)
        Dim nResult83 As Currency, nResult84 As Currency
        
        nResult83 = IIf(vCta83D > vCta83H, vCta83D - vCta83H, vCta83H - vCta83D)
        nResult84 = IIf(vCta84D > vCta84H, vCta84D - vCta84H, vCta84H - vCta84D)
            
        If nResult83 <> nResult84 Then
            vDife8 = 0
            vDife8 = IIf(nResult83 > nResult84, nResult83 - nResult84, nResult84 - nResult83)
            'vDife8 = Abs(nResult83 - nResult84)
            If Abs(nResult83) > Abs(nResult84) Then
                'SI 83 ES MAYOR
                If vCta83D > vCta83H Then
                   
                   InsertaAsientoDN gdHoraGrab, "8321", 0, vDife8, "3", vAgencia
                   
                   InsertaAsientoDN gdHoraGrab, lsCVMEGanacia, vDife8, 0, "3", vAgencia
                   
                    
                Else
                     
                    InsertaAsientoDN gdHoraGrab, "8321", vDife8, 0, "3", vAgencia
                    
                    InsertaAsientoDN gdHoraGrab, lsCVMEGanacia, 0, vDife8, "3", vAgencia
                    
                End If
            Else
                'SI 84 ES MAYOR
                If vCta84D > vCta84H Then
                    
                    InsertaAsientoDN gdHoraGrab, "8321", 0, vDife8, "3", vAgencia
                    
                    InsertaAsientoDN gdHoraGrab, lsCVMEPerdida, vDife8, 0, "3", vAgencia
                    
                    
                Else
                    
                    InsertaAsientoDN gdHoraGrab, "8321", vDife8, 0, "3", vAgencia
                    
                    InsertaAsientoDN gdHoraGrab, lsCVMEPerdida, 0, vDife8, "3", vAgencia
                    
                End If
            End If
        End If
        '**************************************************************************************
        Dim vCta82D As Currency, vCta82H As Currency, vCta81D As Currency, vCta81H As Currency
        'CUADRE DE CTAS DE ORDEN (SILVITA)   82 = 81
        vCta82D = 0: vCta82H = 0
        Set RegTmp = New ADODB.Recordset
       
        Set RegTmp = CuadreCtas(gdHoraGrab, "82%")
        
        If RegTmp Is Nothing Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta82D = vCta82D + RegTmp!nDebe
                vCta82H = vCta82H + RegTmp!nhaber
                RegTmp.MoveNext
            Loop
            RegTmp.Close
        End If
        
        Set RegTmp = Nothing
        vCta82D = Abs(vCta82D)
        vCta82H = Abs(vCta82H)
        
        vCta81D = 0: vCta81H = 0
        Set RegTmp = New ADODB.Recordset
        
        Set RegTmp = CuadreCtas(gdHoraGrab, "81%")
       
        If RegTmp Is Nothing Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta81D = vCta81D + RegTmp!nDebe
                vCta81H = vCta81H + RegTmp!nhaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        vCta81D = Abs(vCta81D)
        vCta81H = Abs(vCta81H)
        
        Dim nResult81 As Currency, nResult82 As Currency
        
        nResult81 = IIf(vCta81D > vCta81H, vCta81D - vCta81H, vCta81H - vCta81D)
        nResult82 = IIf(vCta82D > vCta82H, vCta82D - vCta82H, vCta82H - vCta82D)
        
        'Validación
        If nResult82 <> nResult81 Then
            vDife8 = 0
            vDife8 = IIf(nResult81 > nResult82, nResult81 - nResult82, nResult82 - nResult81)
            'vDife8 = Abs((vCta82D + vCta82H) - (vCta81D + vCta81H))
            If Abs(nResult82) > Abs(nResult81) Then
                'SI 82 MAYOR
                If vCta82D > vCta82H Then
                    
                    InsertaAsientoDN gdHoraGrab, "8221", vDife8, 0, "3", vAgencia
                  
                    InsertaAsientoDN gdHoraGrab, lsCVMEGanacia, vDife8, 0, "3", vAgencia
                    
                Else
                
                    InsertaAsientoDN gdHoraGrab, "8221", vDife8, 0, "3", vAgencia
                    
                    InsertaAsientoDN gdHoraGrab, lsCVMEGanacia, 0, vDife8, "3", vAgencia
                    
                End If
            Else
                'SI 81 MAYOR
                If vCta81D > vCta81H Then
                    
                    
                    InsertaAsientoDN gdHoraGrab, "8221", 0, vDife8, "3", vAgencia
                    
                    InsertaAsientoDN gdHoraGrab, lsCVMEPerdida, vDife8, 0, "3", vAgencia
                   
                Else
                
                   
                    InsertaAsientoDN gdHoraGrab, "8221", vDife8, 0, "3", vAgencia
                    
                    InsertaAsientoDN gdHoraGrab, lsCVMEPerdida, 0, vDife8, "3", vAgencia
                    
                    
                End If
            End If
        End If
        
        '**************************************************************************************
        Dim vCta86D As Currency, vCta86H As Currency, vCta85D As Currency, vCta85H As Currency
        'CUADRE DE CTAS DE ORDEN (SIMON) 86 = 85
        vCta86D = 0: vCta86H = 0
        Set RegTmp = New ADODB.Recordset
        
        Set RegTmp = CuadreCtas(gdHoraGrab, "86%")
        
        If (RegTmp) Is Nothing Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta86D = vCta86D + RegTmp!nDebe
                vCta86H = vCta86H + RegTmp!nhaber
                RegTmp.MoveNext
            Loop
            RegTmp.Close
        End If
       
        Set RegTmp = Nothing
        vCta86D = Abs(vCta86D)
        vCta86H = Abs(vCta86H)
        
        vCta85D = 0: vCta85H = 0
        Set RegTmp = New ADODB.Recordset
        
        Set RegTmp = CuadreCtas(gdHoraGrab, "85%")
       
        If RegTmp Is Nothing Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vCta85D = vCta85D + RegTmp!nDebe
                vCta85H = vCta85H + RegTmp!nhaber
                RegTmp.MoveNext
            Loop
            RegTmp.Close
        End If
        
        Set RegTmp = Nothing
        vCta85D = Abs(vCta85D)
        vCta85H = Abs(vCta85H)
        
        Dim nResult85 As Currency, nResult86 As Currency
        
        nResult85 = IIf(vCta85D > vCta85H, vCta85D - vCta85H, vCta85H - vCta85D)
        nResult86 = IIf(vCta86D > vCta86H, vCta86D - vCta86H, vCta86H - vCta86D)
        
        'Validación
        If nResult86 <> nResult85 Then
            vDife8 = 0
            vDife8 = IIf(nResult85 > nResult86, nResult85 - nResult86, nResult86 - nResult85)
            If Abs(nResult85) > Abs(nResult86) Then
                'SI 85 MAYOR
                If vCta85D > vCta85H Then
                
                    InsertaAsientoDN gdHoraGrab, "8528", 0, vDife8, "3", vAgencia
                    
                    InsertaAsientoDN gdHoraGrab, lsCVMEGanacia, vDife8, 0, "3", vAgencia
  
                Else
                   
                    InsertaAsientoDN gdHoraGrab, "8528", vDife8, 0, "3", vAgencia
                    
                    InsertaAsientoDN gdHoraGrab, lsCVMEGanacia, 0, vDife8, "3", vAgencia
                    
                    
                End If
            Else
                'SI 86 MAYOR
                If vCta86D > vCta86H Then
                    
                    InsertaAsientoDN gdHoraGrab, "8528", 0, vDife8, "3", vAgencia
                    
                    InsertaAsientoDN gdHoraGrab, lsCVMEPerdida, vDife8, 0, "3", vAgencia
                    
                Else
                    
                    InsertaAsientoDN gdHoraGrab, "8528", vDife8, 0, "3", vAgencia
                    
                    InsertaAsientoDN gdHoraGrab, lsCVMEPerdida, 0, vDife8, "3", vAgencia
                    
                End If
            End If
        End If
        
        RegTran.Close
        Set RegTran = Nothing
       
    End If
        
        
        
    'Rutina para obtener asientos de la Compra y Venta de Dolares
    Dim CVDebe As Currency, CVHaber As Currency
    Set RegTmp = New ADODB.Recordset
   
    Set RegTmp = AsientosCompraVentaME(ldFechaAsiento, gAsientoProcesoSiAsientoCVME)
   
    If RegTmp Is Nothing Then
    Else
        With RegTmp
            Do While Not .EOF
                Set RegOpeCta = New ADODB.Recordset
                
                Set RegOpeCta = ObtenerCtaNeg(!cOpeCod)
              
                If RegOpeCta Is Nothing Then
                    lsmenError = lsmenError & " - Código de Operación No Reconocido " & !cOpeCod & Chr$(10)
                    vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & !cOpeCod & " operación compra/venta no reconocida."
                Else
                    Do While Not RegOpeCta.EOF
                        vCodConta = Trim(RegOpeCta!cCtaContCod)
                        vAG = VarAG(Mid(RegTmp!cMovNro, 15, 5), vCodConta)
                        vCodConta = Replace(vCodConta, "AG", vAG, 1, 1, vbTextCompare)
                        vParche = AsientoParche(vCodConta, True)
                        vAgencia = RegTmp!Agencia
                        If Len(vParche) > 0 Then
                            vCodConta = vParche
                        End If
                      '**************************************************************************
                      'Si existe lo actualiza; caso contrario lo agrega
                      'MARCA 12
                        'If Trim(vCodConta) = "19180207" Then Stop
                        
                        If RegOpeCta!cOpeCtaDH = "D" Then
                            If Left(vCodConta, 4) = lsCajaSoles Then
                                'sSql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod) " & _
                                    " VALUES('" & gdHoraGrab & "','" & vCodConta & "'," & Format(!Monto * !nTipCambio, "#0.00") & ",0,'1','" & vAgencia & "'," & !nMovNro & ",'" & !cOpeCod & "' )"
                               
                                InsertaAsientoDN gdHoraGrab, vCodConta, !NMOVSOLES, 0, "1", vAgencia, !cOpeCod, , !nMovNro
                                
                                CVDebe = CVDebe + Format(!NMOVSOLES, "#0.00")
                                
                            ElseIf Left(vCodConta, 4) = lsCajaDolares Then 'Or Left(vCodConta, 4) = "2825"
                                InsertaAsientoDN gdHoraGrab, vCodConta, Round(!Monto * TCF, 2), 0, "1", vAgencia, !cOpeCod, , !nMovNro
                               
                                InsertaAsientoDN gdHoraGrab, vCodConta, !Monto, 0, "2", vAgencia, !cOpeCod, , !nMovNro
 
                                CVDebe = CVDebe + Round(!Monto * TCF, 2)
                            Else
                                lsmenError = lsmenError & " - Cuenta en Compra y Venta no reconocida" & Chr$(10)
                                vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & vCodConta & " cta. de compra/venta no reconocida."
                            End If
                        ElseIf RegOpeCta!cOpeCtaDH = "H" Then
                            If Left(vCodConta, 4) = lsCajaSoles Then
                            
                                InsertaAsientoDN gdHoraGrab, vCodConta, 0, Round(!Monto * !nTipCambio, 2), "1", vAgencia, !cOpeCod, , !nMovNro
                                
                                CVHaber = CVHaber + Round(!Monto * !nTipCambio, 2)
                            ElseIf Left(vCodConta, 4) = lsCajaDolares Then 'Or Left(vCodConta, 4) = "2825"
                                
                                InsertaAsientoDN gdHoraGrab, vCodConta, 0, Round(!Monto * TCF, 2), "1", vAgencia, !cOpeCod, , !nMovNro
                                
                                InsertaAsientoDN gdHoraGrab, vCodConta, 0, (!Monto), "2", vAgencia, !cOpeCod, , !nMovNro
                                                            
                                CVHaber = CVHaber + Round(!Monto * TCF, 2)
                            Else
                                lsmenError = lsmenError & " - Cuenta en Compra y Venta no reconocida" & Chr$(10)
                                vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & vCodConta & " cta. de compra/venta no reconocida."
                            End If
                        Else
                            lsmenError = lsmenError & " - Operación no reconocida al Insertar " & vbCr & _
                            " operación nro.: " & vCodConta & Chr$(10)
                            vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & RegOpeCta!cOpeCtaDH & " debe/haber no reconocido (compra/venta)."
                        End If
                      RegOpeCta.MoveNext
                    Loop
                    RegOpeCta.Close
                    Set RegOpeCta = Nothing
                End If
                .MoveNext
            Loop
        End With
        RegTmp.Close
        Set RegTmp = Nothing
        'Validación de 0.01 por dolares - Silvita
        Dim nCtaDeb1 As Currency, nCtaHab1 As Currency
        Set RegTmp = New ADODB.Recordset
        
        Set RegTmp = ValidarDolares(gdHoraGrab)
        
        
        If (RegTmp) Is Nothing Then
            'No se encuentra ctas en Dolares o Tipo 1
        Else
            vAgencia = RegTmp!cCodAge
            Do While Not RegTmp.EOF
                Do While RegTmp!cCodAge = vAgencia
                    nCtaDeb1 = nCtaDeb1 + RegTmp!nDebe
                    nCtaHab1 = nCtaHab1 + RegTmp!nhaber
                    RegTmp.MoveNext
                    If RegTmp.EOF Then Exit Do
                Loop
                vCtaDolar = nCtaDeb1 - nCtaHab1
                If vCtaDolar > 0 Then
                   
                    InsertaAsientoDN gdHoraGrab, lsCVMEGanacia, 0, Abs(vCtaDolar), "1", vAgencia
                    
                ElseIf vCtaDolar < 0 Then
                
                    InsertaAsientoDN gdHoraGrab, lsCVMEGanacia, Abs(vCtaDolar), 0, "1", vAgencia
        
                End If
                If RegTmp.EOF Then Exit Do
                vAgencia = RegTmp!cCodAge
                nCtaDeb1 = 0: nCtaHab1 = 0
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
    End If
    
    'Asiento para sobrantes y faltantes
    Dim vCodConta1 As String
    Dim vCodConta2 As String
    
    Set RegTmp = New ADODB.Recordset
   
    Set RegTmp = ObtenerSobrantesFaltantes(gAsientoProcesoSiAsientoSofFal, ldFechaAsiento)
 
    If (RegTmp) Is Nothing Then
    Else
        With RegTmp
            Do While Not .EOF
                'Variables para cambio en el CodContable
                ssql = ""
                vMoneda = !nMoneda
                vAgencia = !Agencia
                'If !Agencia = "07" Then Stop
                
                If Not (vMoneda = Moneda.gMonedaNacional Or vMoneda = Moneda.gMonedaExtranjera) Then
                    lsmenError = lsmenError & " - Código Errado, Moneda no definida " & !cCtaCod & Chr$(10)
                    vNoCtaCnt = vNoCtaCnt & Chr$(10) & "     * " & !cCtaCod & " moneda no definida(sobrantes/faltantes)."
                End If
                If !cCodOpe = "901020" Or !cCodOpe = "901010" Then ' Faltante de caja
                    vCodConta1 = "11M102AG"
                    vCodConta2 = IIf(!Monto > 1, "19M80202AG", "63M10909")
                    
                    vParche = AsientoParche(vCodConta2, True)
                   
                    If Len(vParche) > 0 Then
                        vCodConta2 = vParche
                    End If
                    'Cambios en Moneda y Agencia
                    vCodConta1 = Replace(vCodConta1, "M", vMoneda, 1, 1, vbTextCompare)
                    vAG = VarAG(pscodcmac & !Agencia, vCodConta1)
                    vCodConta1 = Trim(Replace(vCodConta1, "AG", vAG, 1, 1, vbTextCompare))
                    vCodConta2 = Replace(vCodConta2, "M", vMoneda, 1, 1, vbTextCompare)
                    vCodConta2 = Trim(Replace(vCodConta2, "AG", !Agencia, 1, 1, vbTextCompare))
         
                    vParche = AsientoParche(vCodConta1, True)
                   
                    If Len(vParche) > 0 Then
                        vCodConta1 = vParche
                    End If
                    
                    'Verifica si el asiento ya fue creado
                    'Si existe lo actualiza; caso contrario lo agrega
                       
                        InsertaAsientoDN gdHoraGrab, vCodConta1, 0, !Monto, "0", vAgencia
                    
                        If Mid(vCodConta1, 3, 1) = Moneda.gMonedaExtranjera Then
                           InsertaAsientoDN gdHoraGrab, vCodConta1, 0, (!Monto * TCF), "3", vAgencia
                        
                        End If
                        
                            InsertaAsientoDN gdHoraGrab, vCodConta2, !Monto, 0, "0", vAgencia
  
                        If Mid(vCodConta2, 3, 1) = Moneda.gMonedaExtranjera Then
                           
                            InsertaAsientoDN gdHoraGrab, vCodConta2, (!Monto * TCF), 0, "3", vAgencia
                            
                        End If
                Else
                    vCodConta1 = "11M102AG"
                    vCodConta2 = IIf(!Monto > 5, "29M20102AG", "52M2290299AG")
                
                    vParche = AsientoParche(vCodConta2, True)
                   
                    If Len(vParche) > 0 Then
                        vCodConta2 = vParche
                    End If
                    'Cambios en Moneda y Agencia
                    vCodConta1 = Replace(vCodConta1, "M", vMoneda, 1, 1, vbTextCompare)
                    vAG = VarAG(pscodcmac & !Agencia, vCodConta1)
                    vCodConta1 = Trim(Replace(vCodConta1, "AG", vAG, 1, 1, vbTextCompare))
                    vCodConta2 = Replace(vCodConta2, "M", vMoneda, 1, 1, vbTextCompare)
                    vCodConta2 = Trim(Replace(vCodConta2, "AG", !Agencia, 1, 1, vbTextCompare))
                    
                    vParche = AsientoParche(vCodConta1, True)
                    
                    If Len(vParche) > 0 Then
                        vCodConta1 = vParche
                    End If
                    
                    'Verifica si el asiento ya fue creado
                    'Si existe lo actualiza; caso contrario lo agrega
                        
                        InsertaAsientoDN gdHoraGrab, vCodConta1, !Monto, 0, "0", vAgencia
                        
                        If Mid(vCodConta1, 3, 1) = Moneda.gMonedaExtranjera Then
                     
                            InsertaAsientoDN gdHoraGrab, vCodConta1, (!Monto * TCF), 0, "3", vAgencia

                        End If
                        
                        InsertaAsientoDN gdHoraGrab, vCodConta2, 0, !Monto, "0", vAgencia
                      
                        If Mid(vCodConta2, 3, 1) = Moneda.gMonedaExtranjera Then
                            
                        InsertaAsientoDN gdHoraGrab, vCodConta2, 0, (!Monto * TCF), "3", vAgencia
                           
                        End If
                End If
                .MoveNext
            Loop
        End With
        RegTmp.Close
        Set RegTmp = Nothing
    End If
    
    '-- Aca
       'Verifica que cuadre DEBE y HABER - Tipo '3'
        Dim vDife As Currency, vDifeSol As Currency
        Dim vMonDebe As Currency, vMonHaber As Currency
        Dim vMonDebeSol As Currency, vMonHaberSol As Currency
        'Carga suma de tipo soles para hallar la diferencia - RAUL
        Set RegTmp = New ADODB.Recordset
      
        Set RegTmp = ObtenerSumaMontoN(gdHoraGrab)
    
        If (RegTmp) Is Nothing Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vMonDebeSol = vMonDebeSol + RegTmp!mondebe
                vMonHaberSol = vMonHaberSol + RegTmp!monhaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        'Halla suma de diferencia de Dolares
        Set RegTmp = New ADODB.Recordset
        Set RegTmp = ObtenerSumaDieferenciaE(gdHoraGrab)
      
        If (RegTmp) Is Nothing Then
            'No se encuentra ctas en Dolares o Tipo 3
        Else
            Do While Not RegTmp.EOF
                vMonDebe = vMonDebe + RegTmp!mondebe
                vMonHaber = vMonHaber + RegTmp!monhaber
                RegTmp.MoveNext
            Loop
        End If
        RegTmp.Close
        Set RegTmp = Nothing
        
        vDifeSol = vMonDebeSol - vMonHaberSol
        
        vMonDebe = Round(vMonDebe, 2)
        vMonHaber = Round(vMonHaber, 2)
        vDife = vMonDebe - vMonHaber
        
        vDife = vDife + vDifeSol
        'Inserta a la cuenta de ganancia (51280401-Haber) o perdida (41180401-Debe)
        'vCtaDolar = (vCta19Debe + vCta28Debe) - (vCta19Haber + vCta28Haber)
        
        vCtaDolar = -1 * vDife 'CtaDolar + vDife
        
        
        If vCtaDolar < 0 Then
            InsertaAsientoDN gdHoraGrab, lsCVMEGanacia & vAgencia, 0, Abs(vCtaDolar), "3", vAgencia
           
        ElseIf vCtaDolar > 0 Then
            InsertaAsientoDN gdHoraGrab, lsCVMEPerdida & vAgencia, Abs(vCtaDolar), 0, "3", vAgencia
   
        End If
    
    'dbCmact.CommitTrans  'Finaliza TRANSACCION
    
    '************** IMPRESION DE ASIENTO ****************
    '****************************************************
    GenerarAsiento = ImpresionAsiento(dFecha, pnaho, pncred, pnpig, pnMonSol, pnMonDol, pdfecsis, psNomCmac, psNomAge, pscoduser, pAsiDia, psmenError, psvobs, psvnoctacnt, pbpasidia, pbcierrerealizado, , lsmenError)
      

End Function


Public Sub TransfiereAsiento(ByVal pdFecha As Date, ByVal pdfecsis As Date, ByRef psmensaje As String, ByVal pscoduser As String, ByVal pscodcmac As String, _
                             Optional ByVal psCodAge As String, Optional psNomAge As String)
    Dim ssql As String
    Dim rs As New ADODB.Recordset
    Dim aAsientoS() As String
    Dim aAsientoD() As String
    Dim aAsientoV() As String
    Dim vdFecha     As Date
    Dim nPos        As Integer
    ReDim aAsientoS(1 To 3, 0 To 0)
    ReDim aAsientoD(1 To 5, 0 To 0)
    ReDim aAsientoV(1 To 5, 0 To 0)
    vdFecha = CDate(pdFecha)
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
   oCon.AbreConexion

   ssql = " SELECT cCtaCnt, cTipo, Round(SUM(nDebe),2) as nDebe, Round(SUM(nHaber),2) as nHaber " _
        & " FROM AsientoDN WHERE datediff(dd,dfecha,'" & Format(vdFecha, "mm/dd/yyyy") & "')= 0 " _
        & " GROUP BY cCtaCnt, cTipo ORDER BY cCtaCnt "
   If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
   rs.CursorLocation = adUseClient
   
   Set rs = oCon.CargaRecordSet(ssql)
   
   rs.ActiveConnection = Nothing
   If rs.EOF Then
      psmensaje = "No se registraron Movimientos en Agencia el día " & vdFecha & "...!"
   Else
      Do While Not rs.EOF

         Select Case rs!cTipo
         Case "0"
            If Mid(rs!cCtaCnt, 3, 1) = "1" Then
               ReDim Preserve aAsientoS(1 To 3, 0 To UBound(aAsientoS, 2) + 1)
               nPos = UBound(aAsientoS, 2)
               aAsientoS(1, nPos) = rs!cCtaCnt
               aAsientoS(2, nPos) = Round(rs!nDebe, 2)
               aAsientoS(3, nPos) = Round(rs!nhaber, 2)
            Else
               nPos = BuscaMatriz(aAsientoD, rs!cCtaCnt, 2)
               If nPos = -1 Then
                  ReDim Preserve aAsientoD(1 To 5, 0 To UBound(aAsientoD, 2) + 1)
                  nPos = UBound(aAsientoD, 2)
               End If
               aAsientoD(1, nPos) = rs!cCtaCnt
               aAsientoD(4, nPos) = Round(rs!nDebe, 2)
               aAsientoD(5, nPos) = Round(rs!nhaber, 2)
            End If
         Case "1"
            nPos = BuscaMatriz(aAsientoV, rs!cCtaCnt, 2)
            If nPos = -1 Then
               ReDim Preserve aAsientoV(1 To 5, 0 To UBound(aAsientoV, 2) + 1)
               nPos = UBound(aAsientoV, 2)
            End If
            aAsientoV(1, nPos) = rs!cCtaCnt
            aAsientoV(2, nPos) = Round(rs!nDebe, 2)
            aAsientoV(3, nPos) = Round(rs!nhaber, 2)
         Case "2"
            nPos = BuscaMatriz(aAsientoV, rs!cCtaCnt, 2)
            If nPos = -1 Then
               ReDim Preserve aAsientoV(1 To 5, 0 To UBound(aAsientoV, 2) + 1)
               nPos = UBound(aAsientoV, 2)
            End If
            aAsientoV(1, nPos) = rs!cCtaCnt
            aAsientoV(4, nPos) = Round(rs!nDebe, 2)
            aAsientoV(5, nPos) = Round(rs!nhaber, 2)
         Case "3"
            nPos = BuscaMatriz(aAsientoD, rs!cCtaCnt, 2)
            If nPos = -1 Then
               ReDim Preserve aAsientoD(1 To 5, 0 To UBound(aAsientoD, 2) + 1)
               nPos = UBound(aAsientoD, 2)
            End If
            aAsientoD(1, nPos) = rs!cCtaCnt
            aAsientoD(2, nPos) = Round(rs!nDebe, 2)
            aAsientoD(3, nPos) = Round(rs!nhaber, 2)
         End Select
         rs.MoveNext
      Loop
      If Not GrabaAsientoMigra(aAsientoS, True, , 1, psmensaje, pdfecsis, pscoduser, pdFecha, pscodcmac, psCodAge, psNomAge) Then
         Exit Sub
      End If
      If Not GrabaAsientoMigra(aAsientoD, False, , 2, psmensaje, pdfecsis, pscoduser, pdFecha, pscodcmac, psCodAge, psNomAge) Then
         Exit Sub
      End If
      GrabaAsientoMigra aAsientoV, False, "Compra - Venta de M.E.", 3, psmensaje, pdfecsis, pscoduser, pdFecha, pscodcmac, psCodAge, psNomAge
   End If
End Sub

Private Function GrabaAsientoMigra(paAsiento() As String, plMN As Boolean, Optional psMsg As String = "", Optional pnTipo As Integer = 0, Optional ByRef psmensaje As String, Optional ByVal pdfecsis As Date, Optional pscoduser As String, Optional pdFecha As Date, _
                                   Optional ByVal pscodcmac As String, Optional ByVal psCodAge As String, Optional psNomAge As String) As Boolean

Dim N As Integer
Dim sCtaCod As String
Dim nItem As Integer
Dim gcOpeCod As String
Dim gcMovNro As String
Dim gcGlosa  As String
Dim rs       As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta
Dim oMov As COMDMov.DCOMMov
Set oMov = New COMDMov.DCOMMov
Dim lnMovNro As Long

Dim lTransActiva As Boolean
Dim sObjetoCod  As String
Dim sIFTpo As String
Dim SCtaIfCod As String
Dim ssql As String
On Error GoTo ErrMigra


Set rs = New ADODB.Recordset

oCon.AbreConexion

GrabaAsientoMigra = False

gcMovNro = Format(pdFecha, "yyyymmdd") & String(6, "0") & pscodcmac & psCodAge & "00XXX" & pnTipo
If Len(gcMovNro) <> 25 Then
   psmensaje = "Error en definición de Agencia...!"
   Exit Function
End If
ssql = "SELECT cMovNro FROM Mov WHERE cMovNro like '" & Mid(gcMovNro, 1, 19) & "__" & Right(gcMovNro, 4) & "' and nMovEstado = " & MovEstado.gMovEstContabMovContable & " And nMovFlag <> " & MovFlag.gMovFlagEliminado
If rs.State = adStateOpen Then rs.Close: Set rs = Nothing

Set rs = oCon.CargaRecordSet(ssql)

If Not rs.EOF Then
   psmensaje = "Ya se realizó la Migración de Asiento de " & psNomAge & Chr$(10) & "Por favor verificar...!"
   Exit Function
End If
If pnTipo = 3 Then
   gcOpeCod = "701108"
Else
   gcOpeCod = "701107"
End If
gcMovNro = oMov.GeneraMovNro(pdfecsis, psCodAge, pscoduser, gcMovNro)
gcGlosa = "Asiento Contable " & psNomAge & " en " & IIf(plMN, "M.N.", "M.E.") & " del " & pdFecha & " " & psMsg
If lTransActiva Then
   oCon.RollbackTrans
End If
lTransActiva = True
oCon.BeginTrans      'Iniciamos Transaccion

'Grabamos Mov
'sSql = "INSERT INTO Mov (cMovNro, cOpeCod, cMovDesc, cMovEstado) VALUES ('" & gcMovNro & "','" & gcOpeCod & "','" & gcGlosa & "','0')"
oMov.InsertaMov gcMovNro, gcOpeCod, gcGlosa

lnMovNro = oMov.GetnMovNro(gcMovNro)

nItem = 0
For N = 1 To UBound(paAsiento, 2)
   'Verificamos la existencia de Cuenta Contable
   sCtaCod = VerificaCuenta(paAsiento(1, N), sObjetoCod, sIFTpo, SCtaIfCod)
   'Grabamos MovCta
   If sCtaCod <> "" Then
      If Val(paAsiento(2, N)) > 0 Then
         nItem = nItem + 1
         oMov.InsertaMovCta lnMovNro, nItem, sCtaCod, paAsiento(2, N)
         If sObjetoCod <> "" Then  'Grabamos MovObj
            oMov.InsertaMovObj lnMovNro, nItem, 1, Format(ObjEntidadesFinancieras, "00")
            oMov.InsertaMovObjIF lnMovNro, nItem, 1, sObjetoCod, sIFTpo, SCtaIfCod
         End If
         If Not plMN Then
            If Val(paAsiento(4, N)) > 0 Then
               oMov.InsertaMovMe lnMovNro, nItem, paAsiento(4, N)
            End If
         End If
      End If
      If Val(paAsiento(3, N)) > 0 Then
         nItem = nItem + 1
         oMov.InsertaMovCta lnMovNro, nItem, sCtaCod, Val(paAsiento(3, N)) * -1
         If sObjetoCod <> "" Then  'Grabamos MovObj
            oMov.InsertaMovObj lnMovNro, nItem, 1, Format(ObjEntidadesFinancieras, "00")
            oMov.InsertaMovObjIF lnMovNro, nItem, 1, sObjetoCod, sIFTpo, SCtaIfCod
         End If
         If Not plMN Then
            If Val(paAsiento(5, N)) > 0 Then
               oMov.InsertaMovMe lnMovNro, nItem, Val(paAsiento(5, N)) * -1
            End If
         End If
      End If
   Else
      psmensaje = "Cuenta Contable " & paAsiento(1, N) & " del día " & pdFecha & " Tipo " & Format(pnTipo, "00") & " no existe. Por favor verificar"
      oCon.RollbackTrans
      lTransActiva = False
      Exit Function
   End If
Next
oCon.CommitTrans
lTransActiva = False
GrabaAsientoMigra = True

Exit Function
ErrMigra:
   If lTransActiva Then
      oCon.RollbackTrans
      lTransActiva = False
   End If
   
End Function


Public Function ImpresionAsiento(ByVal dFecha As Date, ByVal pnaho As Integer, ByVal pncred As Integer, ByVal pnpig As Integer, ByVal pnMonSol As Integer, ByVal pnMonDol As Integer, _
                                 ByVal pdfecsis As Date, ByVal psNomCmac As String, ByVal psNomAge As String, ByVal pscoduser As String, ByVal pAsiDia As String, ByRef psmensaje As String, ByRef psvobs As String, ByRef psvnoctacnt As String, _
                                 ByRef pbpasidia As String, ByRef pbcierrerealizado As String, Optional ByVal lscad As String, Optional ByVal psmenError As String) As String

    Dim vNoCtaCnt As String
    Dim vEst As String
    Dim vcad As String
    Dim x As Long
    Dim vespacio As Long, vLenNomb As Long
    Dim RegTmp As ADODB.Recordset
    Dim tmpSql As String
    Dim TCF As Currency, TCC As Currency, TCV As Currency
    Dim lsCajaDolares As String
    Dim lsCajaSoles As String
    Dim vRTFImp As String
    Dim pLineasMax As Double
    Dim gdHoraGrab As String
    
    Dim oAsi As COMDCajaGeneral.DCOMAsiento
    Set oAsi = New COMDCajaGeneral.DCOMAsiento
    
    Dim oFunI As New COMFunciones.FCOMImpresion
    Dim oFunV As New COMFunciones.FCOMVarPublicas
    
    vLenNomb = 70
    vespacio = vLenNomb + 54
    
    lsCajaSoles = oAsi.GetAsientoParametro(4)
    lsCajaDolares = oAsi.GetAsientoParametro(5)
    
    gdHoraGrab = Format(dFecha & " " & Time, oFunV.gsFormatoFechaHora)
    
    'GENERA IMPRESION DE LOS ASIENTOS
    Dim vIndice As Long  'contador de Item
    Dim vLineas As Long
    Dim vPage As Long
    Dim vMoneAnte As String
    Dim vSumDebe As Currency
    Dim vSumHaber As Currency
    Dim vProduc As String
    Dim vDolar As Boolean
    Dim sSqlDolar As String
    Dim vCabecera As String
    vDolar = False
    Dim vDiaSinOpe  As String
    'VALIDACION DE ASIENTOS
    Dim vObs As String
    Dim lsmenError As String
    vCabecera = "NRO. DE CUENTA                                 DESCRIPCION                                          DEBE          HABER" & Chr$(10)
    Dim rsAge As ADODB.Recordset
    Dim oAre As COMDConstantes.DCOMActualizaDatosArea
    Set rsAge = New ADODB.Recordset
    Set oAre = New COMDConstantes.DCOMActualizaDatosArea
    Set rsAge = oAre.GetAgencias(, False)
    vObs = ""
    Do While Not rsAge.EOF
        'Me.Caption = "Validando asientos de :" & Trim(rsAge!Descripcion)
        
        vObs = vObs & ValidaOk(dFecha, True, rsAge!Codigo)
        rsAge.MoveNext
        DoEvents
    Loop
    'Verifica la veracidad del asiento contable
    'Me.Caption = "Verificando Veracidad de asiento Contable"
    
    lsmenError = lsmenError & " " & psmenError
    If Len(Trim(vObs)) > 0 Or Len(Trim(vNoCtaCnt)) > 0 Then
        lsmenError = lsmenError & " - El asiento presenta Observaciones ! " & Chr$(10)
        vProduc = Chr$(10) & Space(40) & "  A S I E N T O     ¡   N O     V A L I D O   !" & Chr$(10)
        vProduc = vProduc & "   OBSERVACIONES DEL ASIENTO : " & Chr$(10) & _
                            "  ============================="
        If Len(Trim(vNoCtaCnt)) > 0 Then vProduc = vProduc & vNoCtaCnt & Chr$(10)
        If Len(Trim(vObs)) > 0 Then vProduc = vProduc & vObs & Chr$(10)
        vProduc = vProduc & Chr$(10) & Chr$(10)
        vLineas = Repetido(vProduc, Chr$(10))
    Else
        '*************************************************************************************
      
        ActualizaAsientoValida pAsiDia, pscoduser, dFecha, gdHoraGrab
      
       
        '*************************************************************************************
    End If
    'Determina Productos
    vProduc = vProduc & "   DE : " & Chr$(10) & "  ======" & Chr$(10)
    If pnaho = 1 Then vProduc = vProduc & "     *  AHORROS " & Chr$(10)
    If pncred = 1 Then vProduc = vProduc & "     *  CREDITOS " & Chr$(10)
    If pnpig = 1 Then vProduc = vProduc & "     *  CREDITO PIGNORATICIO " & Chr$(10)
   
    
    'Verifica que condiciones se ha de tener en cuenta:  Soles - Dolares
    vEst = "": vcad = ""
    If pnMonSol = 1 Then vEst = vEst & Moneda.gMonedaNacional
    If pnMonDol = 1 Then vEst = vEst & Moneda.gMonedaExtranjera
    For x = 1 To Len(vEst) Step 1
        vcad = vcad & "'" & Mid(vEst, x, 1) & "'"
        If x <> Len(vEst) Then vcad = vcad & ","
    Next x
    ' comentado por avmm pase todo a una funcion
'    If Len(vEst) > 0 Then
'
'    End If
    
    vMoneAnte = ""
    vLineas = vLineas + 12: vPage = 1
    vRTFImp = psNomCmac & Space(vLenNomb + 6) & Format(pdfecsis & " " & Time, "dd/mm/yyyy hh:mm:ss") & Chr$(10)
    vRTFImp = vRTFImp & Space(vLenNomb + 16) & " Página :" & oFunI.ImpreFormat(vPage, 5, 0) & Chr$(10)
    vRTFImp = vRTFImp & oFunI.ImpreFormat(UCase(psNomAge), 25) & Chr$(10)
    vRTFImp = vRTFImp & oFunI.ImpreFormat("   ASIENTO CONTABLE DEL DIA " & Format(dFecha, "dd/mm/yyyy"), 44, 43) & Chr$(10)
    vRTFImp = vRTFImp & oFunI.ImpreFormat(String(40, "="), 44, 42) & Chr$(10)
    If CDate(dFecha) = pdfecsis Then
        
        If Not CierreRealizado(pdfecsis, 1) Then
            vRTFImp = vRTFImp & oFunI.ImpreFormat(" NO VALIDO (ANTES DEL CIERRE DE DIA)", 40, 45) & Chr$(10)
        Else
            'Para los dias que no hay movimientos
           
            Set RegTmp = ObtenerDiasSinMov(gdHoraGrab)
          
            If (RegTmp) Is Nothing Then
                vRTFImp = vRTFImp & vbCr & Space(28) & "N O   S E   R E A L I Z A R O N   O P E R A C I O N E S   E N   E L   D I A"
            Else
                If VerificaDiaHabil(gdHoraGrab, 3) And Not CierreRealizado(2) Then
                    vRTFImp = vRTFImp & oFunI.ImpreFormat(" NO VALIDO (ANTES DEL CIERRE DE MES)", 40, 45) & Chr$(10)
                End If
            End If
            RegTmp.Close
            Set RegTmp = Nothing
        End If
    End If
    '------ solo imprimir la observacion del asiento avmm 09/02/2006 -------------
    Dim lscadObs As String
    lscadObs = vRTFImp & vProduc
    '-----------------------------------------------------------------------------
    If Trim(lsmenError) > 0 Then
        psmensaje = "Aisento Presenta Observaciones .... !"
    End If
    
    Set RegTmp = New ADODB.Recordset
   
    Set RegTmp = ObtenerSumaMNE(0, pnMonSol, pnMonDol, gdHoraGrab, vcad)
   
    If (RegTmp) Is Nothing Then
     
    Else
        With RegTmp
        Do While Not .EOF
            If !Moneda <> vMoneAnte Then
                If !Moneda = Moneda.gMonedaNacional Then
                    vRTFImp = vRTFImp & Chr$(10) & "   MONEDA : SOLES" & Chr$(10)
                    vLineas = vLineas + 1
                Else
                    If pnMonSol = 1 And (vSumDebe > 0 Or vSumHaber > 0) Then
                        vRTFImp = vRTFImp & String(vespacio, "-") & Chr$(10)
                        vRTFImp = vRTFImp & Space(vLenNomb + 12) & oFunI.ImpreFormat("TOTAL", 11, 0) & _
                            oFunI.ImpreFormat(vSumDebe, 12, , True) & oFunI.ImpreFormat(vSumHaber, 12, , True) & Chr$(10)
                        vRTFImp = vRTFImp & String(vespacio, "-") & Chr$(10)
                        vSumDebe = 0: vSumHaber = 0
                        vLineas = vLineas + 3
                    End If
                    vDolar = True
                    If pnMonSol = 1 Then
                        vPage = vPage + 1
                        vRTFImp = vRTFImp & Chr$(12) & Chr$(10)
                        vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & oFunI.ImpreFormat(vPage, 5, 0) & Chr$(10)
                    End If
                    vRTFImp = vRTFImp & Chr$(10) & "   MONEDA : DOLARES" & Chr$(10)
                    vLineas = 5 'vLineas + 2
                End If
                vRTFImp = vRTFImp & String(vespacio, "-") & Chr$(10)
                vRTFImp = vRTFImp & Space(1) & vCabecera
                vRTFImp = vRTFImp & String(vespacio, "-") & Chr$(10)
                vLineas = vLineas + 3
            End If
            vMoneAnte = !Moneda
            vRTFImp = vRTFImp & oFunI.ImpreFormat(!cCtaCnt, 21, 1) & oFunI.ImpreFormat(UCase(CuentaNombre(!cCtaCnt, True)), vLenNomb, 1) & oFunI.ImpreFormat(Round(!nDebe, 2), 12, , True) & oFunI.ImpreFormat(Round(!nhaber, 2), 12, , True) & Chr$(10)
            vSumDebe = vSumDebe + Round(!nDebe, 2)
            vSumHaber = vSumHaber + Round(!nhaber, 2)
            vLineas = vLineas + 1
            If vLineas > pLineasMax Then
                vPage = vPage + 1
                vRTFImp = vRTFImp & Chr$(12) & Chr$(10)
                vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & oFunI.ImpreFormat(vPage, 5, 0) & Chr$(10) & Chr$(10)
                vRTFImp = vRTFImp & String(vespacio, "-") & Chr$(10)
                vRTFImp = vRTFImp & Space(1) & vCabecera
                vRTFImp = vRTFImp & String(vespacio, "-") & Chr$(10)
                vLineas = 5
            End If
            .MoveNext
        Loop
        End With
        vRTFImp = vRTFImp & String(vespacio, "-") & Chr$(10)
        vRTFImp = vRTFImp & Space(vLenNomb + 12) & oFunI.ImpreFormat("TOTAL", 11, 0) & _
            oFunI.ImpreFormat(vSumDebe, 12, , True) & oFunI.ImpreFormat(vSumHaber, 12, , True) & Chr$(10)
        vRTFImp = vRTFImp & String(vespacio, "-") & Chr$(10)
        vLineas = vLineas + 3
        RegTmp.Close
        Set RegTmp = Nothing
        
        If vDolar Then
            Set RegTmp = New ADODB.Recordset
            Set RegTmp = ObtenerSumaMNE(3, pnMonSol, pnMonDol, gdHoraGrab, vcad)
          
            If (RegTmp) Is Nothing Then
               
            Else
                If vLineas + 8 >= pLineasMax Then
                    vPage = vPage + 1
                    vRTFImp = vRTFImp & Chr$(12) & Chr$(10)
                    vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & oFunI.ImpreFormat(vPage, 5, 0) & Chr$(10) & Chr$(10)
                    vLineas = 5
                End If
                vRTFImp = vRTFImp & Chr$(10) & "   CONVERSION DE DOLARES" & Chr$(10)
                vRTFImp = vRTFImp & "     *  TIPO CAMBIO FIJO   : " & oFunI.ImpreFormat(TCF, 6, 3) & Chr$(10)
                vRTFImp = vRTFImp & "     *  TIPO CAMBIO VENTA  : " & oFunI.ImpreFormat(TCV, 6, 3) & Chr$(10)
                vRTFImp = vRTFImp & "     *  TIPO CAMBIO COMPRA : " & oFunI.ImpreFormat(TCC, 6, 3) & Chr$(10)
                vRTFImp = vRTFImp & String(vespacio, "-") & Chr$(10)
                vRTFImp = vRTFImp & Space(1) & vCabecera
                vRTFImp = vRTFImp & String(vespacio, "-") & Chr$(10)
                vLineas = vLineas + 8
                vSumDebe = 0: vSumHaber = 0
                With RegTmp
                Do While Not .EOF
                    vRTFImp = vRTFImp & oFunI.ImpreFormat(!cCtaCnt, 21, 1) & oFunI.ImpreFormat(UCase(CuentaNombre(!cCtaCnt, True)), vLenNomb, 1) & oFunI.ImpreFormat(Round(!nDebe, 2), 12, , True) & oFunI.ImpreFormat(Round(!nhaber, 2), 12, , True) & Chr$(10)
                    vSumDebe = vSumDebe + !nDebe
                    vSumHaber = vSumHaber + !nhaber
                    vLineas = vLineas + 1
                    If vLineas > pLineasMax Then
                        vPage = vPage + 1
                        vRTFImp = vRTFImp & Chr$(12) & Chr$(10)
                        vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & oFunI.ImpreFormat(vPage, 5, 0) & Chr$(10) & Chr$(10)
                        vRTFImp = vRTFImp & String(vespacio, "-") & Chr$(10)
                        vRTFImp = vRTFImp & Space(1) & vCabecera
                        vRTFImp = vRTFImp & String(vespacio, "-") & Chr$(10)
                        vLineas = 5
                    End If
                    .MoveNext
                Loop
                End With
                RegTmp.Close
                Set RegTmp = Nothing
                If vLineas + 3 > pLineasMax Then
                    vPage = vPage + 1
                    vRTFImp = vRTFImp & Chr$(12) & Chr$(10)
                    vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & oFunI.ImpreFormat(vPage, 5, 0) & Chr$(10) & Chr$(10)
                    vLineas = 5
                End If
                vRTFImp = vRTFImp & String(vespacio, "-") & Chr$(10)
                vRTFImp = vRTFImp & Space(vLenNomb + 12) & oFunI.ImpreFormat("TOTAL", 11, 0) & _
                    oFunI.ImpreFormat((vSumDebe), 12, , True) & oFunI.ImpreFormat((vSumHaber), 12, , True) & Chr$(10)
                vRTFImp = vRTFImp & String(vespacio, "-") & Chr$(10)
                vLineas = vLineas + 3
            End If
        End If
    End If
    
    '**************************************************************************************
    'CARGA DE : Compra y Venta en Dolares y en Soles
    'Me.Caption = "Cargando asientos de compra - venta dolares y soles"
    
    Dim ComVta(6, 1) As Currency ' 0-Ctas; 1-Debe;  2-Haber
    Erase ComVta
    Dim ComVtaDolar As String
    ComVtaDolar = ""
    
    Set RegTmp = New ADODB.Recordset
   
    Set RegTmp = CompraVentaMEN(gdHoraGrab)
   
    If (RegTmp) Is Nothing Then
    Else
        If vLineas + 19 >= pLineasMax Then
            vPage = vPage + 1
            vRTFImp = vRTFImp & Chr$(12) & Chr$(10)
            vRTFImp = vRTFImp & Space(vLenNomb + 35) & "Página :" & oFunI.ImpreFormat(vPage, 5, 0) & Chr$(10) & Chr$(10)
            vLineas = 5
        End If
    
        With RegTmp
            Do While Not .EOF
                If !cTipo = "1" Then
                    If Left(!cCtaCnt, 4) = lsCajaDolares Then
                        ComVta(1, 0) = !nDebe
                        ComVta(1, 1) = !nhaber
                    ElseIf Left(!cCtaCnt, 4) = lsCajaSoles Then
                        ComVta(3, 0) = !nDebe
                        ComVta(3, 1) = !nhaber
                    ElseIf Left(!cCtaCnt, 4) = "4128" Then
                        ComVta(4, 0) = !nDebe
                    ElseIf Left(!cCtaCnt, 4) = "5128" Then
                        ComVta(4, 1) = !nhaber
                    Else
                        lsmenError = lsmenError & " - Cuenta no reconocida para Compra y Venta de dolares" & Chr$(10)
                    End If
                Else
                    If Left(!cCtaCnt, 4) = lsCajaDolares Then
                        ComVta(5, 0) = !nDebe
                        ComVta(5, 1) = !nhaber
                    Else
                        lsmenError = lsmenError & " - Cuenta no reconocida para Compra y Venta de dolares" & Chr$(10)
                    End If
                End If
                .MoveNext
            Loop
        End With
        
        ComVtaDolar = "" & Chr$(10)
        ComVtaDolar = ComVtaDolar & "   DE : COMPRA Y VENTA EN DOLARES " & Chr$(10)
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & Chr$(10)
        Dim sCajaME As String, sCajaMN As String
        Dim nDebeAcum As Double, nHaberAcum As Double
        
        nDebeAcum = 0
        nHaberAcum = 0
        RegTmp.MoveFirst
        
        Do While Not RegTmp.EOF
            If RegTmp("cTipo") = "2" Then
                sCajaME = RegTmp("cCtaCnt")
                nDebeAcum = nDebeAcum + RegTmp("nDebe")
                nHaberAcum = nHaberAcum + RegTmp("nHaber")
                ComVtaDolar = ComVtaDolar & oFunI.ImpreFormat(sCajaME, 21, 1) & oFunI.ImpreFormat(UCase(CuentaNombre(sCajaME, True)), vLenNomb, 1) & oFunI.ImpreFormat(Round(RegTmp("nDebe"), 2), 12) & oFunI.ImpreFormat(Round(RegTmp("nHaber"), 2), 12) & Chr$(10)
            End If
            RegTmp.MoveNext
        Loop
        
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & Chr$(10)
        ComVtaDolar = ComVtaDolar & Space(vLenNomb + 12) & oFunI.ImpreFormat("TOTAL", 11, 0) & _
                    oFunI.ImpreFormat(Round(nDebeAcum, 2), 12, , True) & _
                    oFunI.ImpreFormat(Round(nHaberAcum, 2), 12, , True) & Chr$(10)
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & Chr$(10)
        
        ComVtaDolar = ComVtaDolar & Chr$(10) & Chr$(10)
        ComVtaDolar = ComVtaDolar & "   DE : COMPRA Y VENTA EN SOLES " & Chr$(10)
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & Chr$(10)
        
        nDebeAcum = 0
        nHaberAcum = 0
        RegTmp.MoveFirst
        Do While Not RegTmp.EOF
            If RegTmp("cTipo") = "1" Then
                sCajaMN = RegTmp("cCtaCnt")
                nDebeAcum = nDebeAcum + RegTmp("nDebe")
                nHaberAcum = nHaberAcum + RegTmp("nHaber")
                ComVtaDolar = ComVtaDolar & oFunI.ImpreFormat(sCajaMN, 21, 1) & oFunI.ImpreFormat(UCase(CuentaNombre(sCajaMN, True)), vLenNomb, 1) & oFunI.ImpreFormat(RegTmp("nDebe"), 12) & oFunI.ImpreFormat(RegTmp("nHaber"), 12) & Chr$(10)
            End If
            RegTmp.MoveNext
        Loop

        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & Chr$(10)
        ComVtaDolar = ComVtaDolar & Space(vLenNomb + 12) & oFunI.ImpreFormat("TOTAL", 11, 0) & _
                    oFunI.ImpreFormat(nDebeAcum, 12, , True) & _
                    oFunI.ImpreFormat(nHaberAcum, 12, , True) & Chr$(10)
        ComVtaDolar = ComVtaDolar & String(vespacio, "-") & Chr$(10)
        ComVtaDolar = ComVtaDolar & Chr$(10)
    End If
    RegTmp.Close
    Set RegTmp = Nothing
    
    'Me.Caption = "Asiento finalizado"
    
    'Envia Asiento al Previo
    If Trim(lsmenError) <> "" Then
        ImpresionAsiento = lscadObs
    Else
        vRTFImp = vRTFImp & ComVtaDolar
        ImpresionAsiento = vRTFImp
    End If
    
    If Len(Trim(ImpresionAsiento)) = 0 Then
        psmensaje = " No existe ningún asiento generado en este día "
    End If
    psvobs = vObs
    psvnoctacnt = vNoCtaCnt
    pbpasidia = pAsiDia
    pbcierrerealizado = CierreRealizado(1)
    Exit Function
End Function

Public Function ExisTipPer(ByVal pCodOpe As String, pConceptoCod As Long, Optional ByVal pNuePlan As Boolean = False) As Boolean
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    
    tmpSql = "SELECT nPersoneria FROM OpeCtaNeg WHERE cOpeCod = '" & pCodOpe & "' And nConcepto = " & pConceptoCod
    Set tmpReg = oCon.CargaRecordSet(tmpSql)
    If (tmpReg.BOF Or tmpReg.EOF) Then
        ExisTipPer = False
    Else
        ExisTipPer = True
        Do While Not tmpReg.EOF
            If Trim(tmpReg!nPersoneria) = "0" Then
                ExisTipPer = False
                Exit Do
            End If
            tmpReg.MoveNext
        Loop
    End If
    tmpReg.Close
    Set tmpReg = Nothing
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Public Function ClienteTipoCTS(ByVal pCuenta As String) As String
    Dim tmpReg As ADODB.Recordset
    Dim tmpSql As String
    Set tmpReg = New ADODB.Recordset
    Dim oCon As COMConecta.DCOMConecta
    Set oCon = New COMConecta.DCOMConecta
    
    oCon.AbreConexion
    
    tmpReg.CursorLocation = adUseClient
    tmpSql = "SELECT cCodInst FROM CaptacCTS WHERE cCtaCod = '" & pCuenta & "'"
    Set tmpReg = oCon.CargaRecordSet(tmpSql)
    Set tmpReg.ActiveConnection = Nothing
    If (tmpReg.BOF Or tmpReg.EOF) Then
        ClienteTipoCTS = ""
    Else
        If Trim(tmpReg!cCodInst) = gConstPersCodCMACT Then
            ClienteTipoCTS = "02"               'Empleado
        Else
            ClienteTipoCTS = "01"               'Cliente
        End If
    End If
    tmpReg.Close
    Set tmpReg = Nothing
End Function


Public Function ValidaOk(ByVal pFecha As Date, Optional pNuePlan As Boolean = False, Optional psAgeCod As String = "", Optional psCodAge As String) As String
Dim vDifere As Currency, vMovAsi As Currency, vMovSal As Currency
Dim vSHoy As Currency, vSAyer As Currency
Dim vDHoy As Currency, vDAyer As Currency
Dim vFecha As Date
Dim nDias As Integer
Dim x As Integer
Dim lsCtaCaja As String
Dim oCon As COMConecta.DCOMConecta
Dim ssql As String
nDias = 30 'Val(ReadVarSis("ADM", "nDiaValAsiento"))
x = 0
ValidaOk = "" ':  vCieDia = 0
'********************************************************************************************
'CTA. 111103AG   --->  A  Caja
vFecha = pFecha
'vSHoy = Billetaje(vFecha, "1", pFecha)
'vDHoy = Billetaje(vFecha, "2", pFecha)

vSHoy = Billetaje(vFecha, "1", psAgeCod)
vDHoy = Billetaje(vFecha, "2", psAgeCod)
If vSHoy = 0 And vDHoy = 0 Then
    vSAyer = 0
    vDAyer = 0
Else
'    Do While (vSAyer = 0 And vDAyer = 0) And nDias > X
'        X = X + 1
'        vFecha = DateAdd("d", -1, vFecha)
'        vSAyer = Billetaje(vFecha, "1", pFecha)
'        vDAyer = Billetaje(vFecha, "2", pFecha)
'    Loop
    vFecha = DateAdd("d", -1, vFecha)
    vSAyer = Billetaje(vFecha, "1", psAgeCod)
    vDAyer = Billetaje(vFecha, "2", psAgeCod)

End If
'Verifica que cuadre la cta. 111103AG o 111102AG (NuevoPlan) (Nela 17.09.2001)
vMovSal = vSHoy - vSAyer
vMovAsi = CtaAsiento(pFecha, "A", "1", "1", pNuePlan, psAgeCod)
vDifere = (vMovSal - vMovAsi)
lsCtaCaja = AsientoParche("111102" & Right(psCodAge, 2), True)
If lsCtaCaja = "" Then
    lsCtaCaja = "111102" & Right(psCodAge, 2)
End If

Set oCon = New COMConecta.DCOMConecta
oCon.AbreConexion

If (vDifere < 0 And vDifere >= -0.05) Then
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
           " VALUES('" & Format(vFecha, "yyyymmdd hh:mm:ss") & "','63110909'," & Abs(vDifere) & ",0,'0','" & psAgeCod & "') "
        oCon.Ejecutar ssql
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
           " VALUES('" & Format(vFecha, "yyyymmdd hh:mm:ss") & "','" & lsCtaCaja & "' ,0," & Abs(vDifere) & ",'0','" & psAgeCod & "') "
        oCon.Ejecutar ssql
ElseIf (vDifere > 0 And vDifere <= 0.05) Then
      ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
           " VALUES('" & Format(vFecha, "yyyymmdd hh:mm:ss") & "','5212290299" & Right(psCodAge, 2) & "' ,0," & Abs(vDifere) & ",'0','" & psAgeCod & "')"
      oCon.Ejecutar ssql
      ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
           " VALUES('" & Format(vFecha, "yyyymmdd hh:mm:ss") & "','" & lsCtaCaja & "' ," & Abs(vDifere) & ",0,'0','" & psAgeCod & "') "
      oCon.Ejecutar ssql
End If
'**
If Abs(vDifere) > 0.05 Then
' VALIDACIONES PARA CTA 11 (SILVITA - NELA) ?????
    If pNuePlan Then
        ValidaOk = ValidaOk & Chr$(10) & "     * No cuadra la cta.cnt. 111102AG en Soles, diferencia de " & Str(vDifere)
    Else
        ValidaOk = ValidaOk & Chr$(10) & "     * No cuadra la cta.cnt. 111103AG en Soles, diferencia de " & Str(vDifere)
    End If
    ValidaOk = ValidaOk & Chr$(10) & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk = ValidaOk & Chr$(10) & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If

vMovSal = vDHoy - vDAyer
vMovAsi = CtaAsiento(pFecha, "A", "2", "2", pNuePlan, psAgeCod)
vDifere = (vMovSal - vMovAsi)
lsCtaCaja = AsientoParche("112102" & Right(psCodAge, 2), True)
If lsCtaCaja = "" Then
    lsCtaCaja = "112102" & Right(psCodAge, 2)
End If

'**
If (vDifere < 0 And vDifere >= -0.05) Then
      ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo) " & _
           " VALUES('" & Format(vFecha, "yyyymmdd hh:mm:ss") & "','63210909'," & Abs(vDifere) & ",0,'0')"
      oCon.Ejecutar ssql
      ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo) " & _
           " VALUES('" & Format(vFecha, "yyyymmdd hh:mm:ss") & "','" & lsCtaCaja & "' ,0," & Abs(vDifere) & ",'0')"
      oCon.Ejecutar ssql
ElseIf (vDifere > 0 And vDifere <= 0.05) Then
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo) " & _
           " VALUES('" & Format(vFecha, "yyyymmdd hh:mm:ss") & "','5222290299" & Right(psCodAge, 2) & " ',0," & Abs(vDifere) & ",'0')"
        oCon.Ejecutar ssql
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo) " & _
           " VALUES('" & Format(vFecha, "yyyymmdd hh:mm:ss") & "','" & lsCtaCaja & "'," & Abs(vDifere) & ",0,'0')"
        oCon.Ejecutar ssql
End If
'**
oCon.CierraConexion
Set oCon = Nothing
If Abs(vDifere) > 0.05 Then
    If pNuePlan Then
        ValidaOk = ValidaOk & Chr$(10) & "     * No cuadra la cta.cnt. 112102AG en Dólares, diferencia de " & Str(vDifere)
    Else
        ValidaOk = ValidaOk & Chr$(10) & "     * No cuadra la cta.cnt. 112103AG en Dólares, diferencia de " & Str(vDifere)
    End If
    ValidaOk = ValidaOk & Chr$(10) & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk = ValidaOk & Chr$(10) & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
'********************************************************************************************
'CTA. 14 - (1419)   ---> B   Créditos
'Verifica que cuadre la cta. 14
vMovSal = Abs(EstadCred(pFecha, "1", 1, psAgeCod) + EstadCred(pFecha, "1", 2, psAgeCod))
vMovAsi = Abs(CtaAsiento(pFecha, "B", "1", , pNuePlan, psAgeCod))
vDifere = (vMovSal - vMovAsi)
If vDifere <> 0 Then
    ValidaOk = ValidaOk & Chr$(10) & "     * No cuadra la cta.cnt. 14 en Soles, diferencia de " & Str(vDifere)
    ValidaOk = ValidaOk & Chr$(10) & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk = ValidaOk & Chr$(10) & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
vMovSal = Abs(EstadCred(pFecha, "2", 1, psAgeCod))
vMovAsi = Abs(CtaAsiento(pFecha, "B", "2", , pNuePlan, psAgeCod))
vDifere = (vMovSal - vMovAsi)
If vDifere <> 0 Then
    ValidaOk = ValidaOk & Chr$(10) & "     * No cuadra la cta.cnt. 14 en Dólares, diferencia de " & Str(vDifere)
    ValidaOk = ValidaOk & Chr$(10) & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk = ValidaOk & Chr$(10) & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
'********************************************************************************************
'CTA. 23,24 y 26  ---> C  Ahorros
'Verifica que cuadren las ctas. 23,24 y 26 o 21 (NuevoPlan)
vMovSal = Abs(EstadAho(pFecha, "1", 1, psAgeCod) + EstadAho(pFecha, "1", 2, psAgeCod) + EstadAho(pFecha, "1", 3, psAgeCod))
vMovAsi = Abs(CtaAsiento(pFecha, "C", "1", , pNuePlan, psAgeCod))
vDifere = Round((vMovSal - vMovAsi), 2)
If vDifere <> 0 Then
    If pNuePlan Then
        ValidaOk = ValidaOk & Chr$(10) & "     * No cuadran las ctas.cnts. 2112, 2113, 2312 y 2313 en Soles, diferencia de " & Str(vDifere)
    Else
        ValidaOk = ValidaOk & Chr$(10) & "     * No cuadran las ctas.cnts. 23, 24 y 26 en Soles, diferencia de " & Str(vDifere)
    End If
    ValidaOk = ValidaOk & Chr$(10) & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk = ValidaOk & Chr$(10) & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
vMovSal = Abs(EstadAho(pFecha, "2", 1, psAgeCod) + EstadAho(pFecha, "2", 2, psAgeCod) + EstadAho(pFecha, "2", 3, psAgeCod))
vMovAsi = Abs(CtaAsiento(pFecha, "C", "2", , pNuePlan, psAgeCod))
vDifere = (vMovSal - vMovAsi)
If vDifere <> 0 Then
    If pNuePlan Then
        ValidaOk = ValidaOk & Chr$(10) & "     * No cuadran las ctas.cnts. 2122, 2123, 2322 y 2323 en Dólares, diferencia de " & Str(vDifere)
    Else
        ValidaOk = ValidaOk & Chr$(10) & "     * No cuadran las ctas.cnts. 23, 24 y 26 en Dólares, diferencia de " & Str(vDifere)
    End If
    ValidaOk = ValidaOk & Chr$(10) & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk = ValidaOk & Chr$(10) & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
If Not ValidaOk = "" And Not psAgeCod = "" Then
    ValidaOk = Chr(10) & Chr(10) & "AGENCIA " & psAgeCod & ": " & ValidaOk
End If
End Function


'****************************** OTRAS FUNCIONES ******************************************
'*****************************************************************************************

Public Function VerificaDiaHabil(ByVal pdFecha As Date, pnTipo As Integer) As Boolean
        Dim lsFchPro As String
        Dim ldFchPro As Date
        Dim lnMes As Integer
        Dim ldAno As Date
        Dim oFunC As New COMFunciones.FCOMCadenas
        ldAno = pdFecha
        lnMes = Month(pdFecha)
        If lnMes = 12 Then
            lnMes = 0
            ldAno = DateAdd("yyyy", 1, ldAno)
        End If
    Select Case pnTipo
        Case 1
        '   Busca Primer día del mes
            lsFchPro = "01" & "/" & oFunC.FillNum(Trim(Str(lnMes + 1)), 2, "0") & "/" & Trim(Str(Year(ldAno)))
            ldFchPro = CDate(lsFchPro)
        Case 2
        '   Busca día anterior al último día del mes para procesar descuento por Inactivas
            lsFchPro = "01" & "/" & oFunC.FillNum(Trim(Str(lnMes + 1)), 2, "0") & "/" & Trim(Str(Year(ldAno)))
            ldFchPro = DateAdd("d", -2, CDate(lsFchPro))
        Case 3
        '   Busca ultimo día del mes para procesar Cierre de Mes
            lsFchPro = "01" & "/" & oFunC.FillNum(Trim(Str(lnMes + 1)), 2, "0") & "/" & Trim(Str(Year(ldAno)))
            ldFchPro = DateAdd("d", -1, CDate(lsFchPro))
    End Select
    If pdFecha = ldFchPro Then
        VerificaDiaHabil = True
    Else
        VerificaDiaHabil = False
    End If
End Function

Private Function BuscaMatriz(paM() As String, psDato As String, Optional pnDimen As Integer = 1) As Integer
Dim N As Integer
BuscaMatriz = -1
For N = LBound(paM, pnDimen) To UBound(paM, pnDimen)
   If paM(1, N) = psDato Then
      BuscaMatriz = N
      Exit Function
   End If
Next
End Function

