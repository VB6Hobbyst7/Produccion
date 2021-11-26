Attribute VB_Name = "gFunAsientoD"
Option Explicit
Dim ssql As String, vCad As String
Public gsCtaCodFoncodes As String

'**DAOR 20100623 **********************************************************
Dim sTC As String, sCD As String, sIK As String, sSC As String, sTpoCredCod As String, sTpoProdCod As String
Dim sNS As String, sIF As String
Dim sTL As String
Dim nTpoInstCorp As Integer
'***************************************************************************

Public Function ValidaOk(ByVal pFecha As Date, Optional pNuePlan As Boolean = False, Optional psAgeCod As String = "") As String
Dim vDifere As Currency, vMovAsi As Currency, vMovSal As Currency
Dim vSHoy As Currency, vSAyer As Currency
Dim vDHoy As Currency, vDAyer As Currency
Dim vFecha As Date
Dim nDias As Integer
Dim x As Integer
Dim lsCtaCaja As String
Dim oCon As DConecta
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
Set oCon = New DConecta
oCon.AbreConexion

vMovSal = vSHoy - vSAyer
vMovAsi = CtaAsiento(pFecha, "A", "1", "1", pNuePlan, psAgeCod, oCon.ConexionActiva)
vDifere = (vMovSal - vMovAsi)
lsCtaCaja = AsientoParche("111102" & Right(gsCodAge, 2), True, oCon.ConexionActiva)
If lsCtaCaja = "" Then
    lsCtaCaja = "111102" & Right(gsCodAge, 2)
End If



If (vDifere < 0 And vDifere >= -0.05) Then
        'ARCV 31-03-2007
        'sSql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','63110909'," & Abs(vDifere) & ",0,'0','" & psAgeCod & "') "
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','421229" & Right(gsCodAge, 2) & "'," & Abs(vDifere) & ",0,'0','" & psAgeCod & "') "
        '-------------
        oCon.ejecutar ssql
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','" & lsCtaCaja & "' ,0," & Abs(vDifere) & ",'0','" & psAgeCod & "') "
        oCon.ejecutar ssql
ElseIf (vDifere > 0 And vDifere <= 0.05) Then
      'ARCV 31-03-2007
      'ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','5212290299" & Right(gsCodAge, 2) & "' ,0," & Abs(vDifere) & ",'0','" & psAgeCod & "')"
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','521229" & Right(gsCodAge, 2) & "' ,0," & Abs(vDifere) & ",'0','" & psAgeCod & "')"
        oCon.ejecutar ssql
      '---------------
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
            " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','" & lsCtaCaja & "' ," & Abs(vDifere) & ",0,'0','" & psAgeCod & "') "
        oCon.ejecutar ssql
End If
'**
If Abs(vDifere) > 0.05 Then
' VALIDACIONES PARA CTA 11 (SILVITA - NELA) ?????
    If pNuePlan Then
        ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "     * No cuadra la cta.cnt. 111102AG en Soles, diferencia de " & Str(vDifere)
    Else
        ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "     * No cuadra la cta.cnt. 111103AG en Soles, diferencia de " & Str(vDifere)
    End If
    ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If

vMovSal = vDHoy - vDAyer
vMovAsi = CtaAsiento(pFecha, "A", "2", "2", pNuePlan, psAgeCod, oCon.ConexionActiva)
vDifere = (vMovSal - vMovAsi)
lsCtaCaja = AsientoParche("112102" & Right(gsCodAge, 2), True, oCon.ConexionActiva)
If lsCtaCaja = "" Then
    lsCtaCaja = "112102" & Right(gsCodAge, 2)
End If

'**
If (vDifere < 0 And vDifere >= -0.05) Then
      'ARCV 31-03-2007
      'sSql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','63210909'," & Abs(vDifere) & ",0,'0')"
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','422229" & Right(gsCodAge, 2) & "'," & Abs(vDifere) & ",0,'0')"
        '--------
        oCon.ejecutar ssql
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','" & lsCtaCaja & "' ,0," & Abs(vDifere) & ",'0')"
        oCon.ejecutar ssql
ElseIf (vDifere > 0 And vDifere <= 0.05) Then
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','522229" & Right(gsCodAge, 2) & " ',0," & Abs(vDifere) & ",'0')"
        oCon.ejecutar ssql
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','" & lsCtaCaja & "'," & Abs(vDifere) & ",0,'0')"
        oCon.ejecutar ssql
End If
'**
If Abs(vDifere) > 0.05 Then
    If pNuePlan Then
        ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "     * No cuadra la cta.cnt. 112102AG en Dólares, diferencia de " & Str(vDifere)
    Else
        ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "     * No cuadra la cta.cnt. 112103AG en Dólares, diferencia de " & Str(vDifere)
    End If
    ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
'********************************************************************************************
'CTA. 14 - (1419)   ---> B   Créditos
'Verifica que cuadre la cta. 14
vMovSal = Abs(EstadCred(pFecha, "1", 1, psAgeCod) + EstadCred(pFecha, "1", 2, psAgeCod))
vMovAsi = Abs(CtaAsiento(pFecha, "B", "1", , pNuePlan, psAgeCod, oCon.ConexionActiva))
vDifere = (vMovSal - vMovAsi)
If vDifere <> 0 Then
    ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "     * No cuadra la cta.cnt. 14 en Soles, diferencia de " & Str(vDifere)
    ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
vMovSal = Abs(EstadCred(pFecha, "2", 1, psAgeCod))
vMovAsi = Abs(CtaAsiento(pFecha, "B", "2", , pNuePlan, psAgeCod, oCon.ConexionActiva))
vDifere = (vMovSal - vMovAsi)
If vDifere <> 0 Then
    ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "     * No cuadra la cta.cnt. 14 en Dólares, diferencia de " & Str(vDifere)
    ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
'********************************************************************************************
'CTA. 23,24 y 26  ---> C  Ahorros
'Verifica que cuadren las ctas. 23,24 y 26 o 21 (NuevoPlan)
vMovSal = Abs(EstadAho(pFecha, "1", 1, psAgeCod) + EstadAho(pFecha, "1", 2, psAgeCod) + EstadAho(pFecha, "1", 3, psAgeCod))
vMovAsi = Abs(CtaAsiento(pFecha, "C", "1", , pNuePlan, psAgeCod, oCon.ConexionActiva))
vDifere = Round((vMovSal - vMovAsi), 2)
If vDifere <> 0 Then
    If pNuePlan Then
        ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "     * No cuadran las ctas.cnts. 2112, 2113, 2312 y 2313 en Soles, diferencia de " & Str(vDifere)
    Else
        ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "     * No cuadran las ctas.cnts. 23, 24 y 26 en Soles, diferencia de " & Str(vDifere)
    End If
    ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
vMovSal = Abs(EstadAho(pFecha, "2", 1, psAgeCod) + EstadAho(pFecha, "2", 2, psAgeCod) + EstadAho(pFecha, "2", 3, psAgeCod))
vMovAsi = Abs(CtaAsiento(pFecha, "C", "2", , pNuePlan, psAgeCod, oCon.ConexionActiva))
vDifere = (vMovSal - vMovAsi)
If vDifere <> 0 Then
    If pNuePlan Then
        ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "     * No cuadran las ctas.cnts. 2122, 2123, 2322 y 2323 en Dólares, diferencia de " & Str(vDifere)
    Else
        ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "     * No cuadran las ctas.cnts. 23, 24 y 26 en Dólares, diferencia de " & Str(vDifere)
    End If
    ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk = ValidaOk & oImpresora.gPrnSaltoLinea & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
If Not ValidaOk = "" And Not psAgeCod = "" Then
    ValidaOk = Chr(10) & Chr(10) & "AGENCIA " & psAgeCod & ": " & ValidaOk
End If
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Function Billetaje(ByVal pFecha As Date, ByVal pMoneda As String, Optional psAgeCod As String = "") As Currency
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
'    tmpSql = " Select Sum(nMonto) Campo From Mov M" _
'           & " Inner Join MovUserEfectivo ME On M.nMovNro = ME.nMovNro" _
'           & " Where M.cMovNro Like '" & Format(pFecha, gsFormatoMovFecha) & "%' And ME.nMonto <> 0 " _
'           & " And ME.cEfectivoCod LIKE '" & pMoneda & "%' And M.cOpeCod IN ('901007','901016') " _
'           & "AND Substring(cMovNro, 18,2) in (SELECT distinct right(cCtaCnt,2) FROM AsientoDN " _
'           & "WHERE datediff(dd,dfecha,'" & Format(pFechaMov, "mm/dd/yyyy") & "') = 0 " _
'           & "AND Substring(cCtaCnt,1,6) IN ('11" & pMoneda & "102','11" & pMoneda & "103'))"

'ARCV 28-03-2007
'tmpSql = "Select ISNULL(SUM(E.nMonto),0) Campo From Mov M INNER JOIN MovUserEfectivo E ON M.nMovNro = E.nMovNro " _
    & "Where M.nMovFlag NOT IN (2,1)  And LEFT(M.cMovNro,8) IN (Select MAX(LEFT(M1.cMovNro,8)) From Mov M1 " _
    & "JOIN MovUserEfectivo E1 ON M1.nMovNro = E1.nMovNro Where M1.nMovFlag NOT IN (2,1) And E1.cEfectivoCod LIKE '" & pMoneda & "%' And " _
    & "LEFT(M1.cMovNro,8) <= '" & Format(pFecha, gsFormatoMovFecha) & "' AND M.cOpeCod in ('901007', '901016')) AND M.cOpeCod in ('901007', '901016') And E.cEfectivoCod LIKE '" & pMoneda & "%'"
 
tmpSql = "SELECT ISNULL(SUM(E.nMonto),0) Campo From Mov M INNER JOIN MovUserEfectivo E ON M.nMovNro = E.nMovNro " _
    & "Where M.nMovFlag NOT IN (2,1)  And LEFT(M.cMovNro,8) IN (Select MAX(LEFT(M1.cMovNro,8)) From Mov M1 " _
    & "JOIN MovUserEfectivo E1 ON M1.nMovNro = E1.nMovNro Where M1.nMovFlag NOT IN (2,1) And E1.cEfectivoCod LIKE '" & pMoneda & "%' And " _
    & "LEFT(M1.cMovNro,8) <= '" & Format(pFecha, gsFormatoMovFecha) & "' AND M.cOpeCod in ('901007', '901016')" _
    & IIf(psAgeCod = "", "", " AND Substring(M1.cMovNro,18,2) = '" & psAgeCod & "'") & ") AND M.cOpeCod in ('901007', '901016') And E.cEfectivoCod LIKE '" & pMoneda & "%'"
'----------

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
Dim oCon As DConecta
Set oCon = New DConecta

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
Dim oCon As DConecta
Set oCon = New DConecta

oCon.AbreConexion

If pOption = 1 Then
    tmpSql = " SELECT (ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(DateAdd("d", -1, pFecha), gsFormatoFecha) & "') = 0 then nSaldo End)),0)" _
           & "         -  ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(pFecha, gsFormatoFecha) & "') = 0 then nSaldo End)),0)) AS Campo" _
           & " FROM CapEstadSaldo WHERE nMoneda = '" & pMoneda & "' And nProducto = " & Producto.gCapAhorros & _
        IIf(psAgeCod = "", "", " and cCodAge = '" & psAgeCod & "'")
ElseIf pOption = 2 Then
    tmpSql = " SELECT    (ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(DateAdd("d", -1, pFecha), gsFormatoFecha) & "') = 0 then nSaldo End)),0)" _
           & "         -  ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(pFecha, gsFormatoFecha) & "') = 0 then nSaldo End)),0)) AS Campo" _
           & " FROM CapEstadSaldo WHERE nMoneda = '" & pMoneda & "' And nProducto = " & Producto.gCapPlazoFijo & _
        IIf(psAgeCod = "", "", " and cCodAge = '" & psAgeCod & "'")
Else
    tmpSql = " SELECT    (ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(DateAdd("d", -1, pFecha), gsFormatoFecha) & "') = 0 then nSaldo End)),0)" _
           & "         -  ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(pFecha, gsFormatoFecha) & "') = 0 then nSaldo End)),0)) AS Campo" _
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

Public Function CtaAsiento(ByVal pFecha As Date, ByVal pTipo As String, ByVal pMoneda As String, _
Optional ByVal pOtroTipo As String = "", Optional ByVal pNuePlan As Boolean = False, Optional psAgeCod As String = "", Optional oCon As ADODB.Connection) As Currency
    Dim tmpReg As ADODB.Recordset
    Dim tmpSql As String, ssql As String
    Set tmpReg = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    'oCon.AbreConexion
    
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
            'ALPA 20110105
'            tmpSql = tmpSql & " AND substring(cCtaCnt,1,3) IN ('14" & pMoneda & "') "
'            tmpSql = tmpSql & "     AND ( substring(cCtaCnt,1,4) Not IN ('1419','1429','1418','1428') "
'            tmpSql = tmpSql & " ) "
'            tmpSql = tmpSql & IIf(psAgeCod <> "", " and SubString(cCtaCod,4,2) = '" & psAgeCod & "' ", "")
            
            tmpSql = tmpSql & " AND substring(cCtaCnt,1,3) IN ('14" & pMoneda & "') "
            tmpSql = tmpSql & "     AND ( substring(cCtaCnt,1,4) Not IN ('1419','1429','1418','1428') and cOpeCod not in ('700104') "
            tmpSql = tmpSql & " ) "
            tmpSql = tmpSql & IIf(psAgeCod <> "", " and right(cCtaCnt,2)  = '" & psAgeCod & "' ", "")
            'tmpSql = tmpSql & IIf(psAgeCod <> "", " and case when cOpeCod in ('100911','100912') then cCodAge else SubString(cCtaCod,4,2) end  = '" & psAgeCod & "' ", "")

        ElseIf pTipo = "C" Then     ' Ahorros
            'Se agrega la 2117 y 2127 por las inmovilizadas forman parte del saldo
            tmpSql = tmpSql & " AND substring(cCtaCnt,1,4) IN ('21" & pMoneda & "2','21" & pMoneda & "3','23" & pMoneda & "2','23" & pMoneda & "3','21" & pMoneda & "7') and cOpeCod not in ('700104')  "
            tmpSql = tmpSql & IIf(psAgeCod <> "", " and SubString(cCtaCod,4,2) = '" & psAgeCod & "' ", "")
        Else
            MsgBox "Tipo en CtaAsiento no Reconocido", vbInformation, " Aviso "
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
        Else
            MsgBox "Tipo en CtaAsiento no Reconocido", vbInformation, " Aviso "
        End If
    End If
    tmpSql = tmpSql & IIf(Len(Trim(pOtroTipo)) = 0, " And cTipo IN ('0')", " Where cTipo IN ('0','" & pOtroTipo & "')")
    'Set tmpReg = oCon.CargaRecordSet(tmpSql)
    'tmpReg.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
    '**ALPA***20080906
    Set tmpReg = oCon.Execute(tmpSql)
    'Set tmpReg.ActiveConnection = Nothing
    If (tmpReg.BOF Or tmpReg.EOF) Then
        CtaAsiento = 0
    Else
        CtaAsiento = IIf(IsNull(tmpReg!Campo), 0, tmpReg!Campo)
    End If
    tmpReg.Close
    Set tmpReg = Nothing
End Function
'********************************************************************************************
'********************************************************************************************
'********************************************************************************************
'Funciones comunes de los Asientos

'Función que retorna los códigos de operación que no deben de tomarse al
' leer el TranDiaria o TranDiariaConsol para la generación de los ASIENTOS
Public Function GetAsiParam(ByVal pTipoAsiento As String, oCon As ADODB.Connection) As String
    Dim tmpReg As ADODB.Recordset
    Dim tmpSql As String
    Dim vCampo As String
    Dim sTip1 As String, sTip2 As String, sTip3 As String, sTip4 As String, sTip5 As String, sTip6 As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    vCampo = IIf(pTipoAsiento = "1", "t.cCodOpe", "t.cCodOpe2")
    tmpSql = "SELECT cCodOpe, cNivOpe FROM AsientoParam  " & _
        " WHERE cTipAsi = '" & pTipoAsiento & "' ORDER BY cTipAsi, cNivOpe "
    Set tmpReg = New ADODB.Recordset
    'tmpReg.CursorLocation = adUseClient
    
    'oCon.AbreConexion
    'Set tmpReg = oCon.CargaRecordSet(tmpSql)
    'ALPA
    Set tmpReg = oCon.Execute(tmpSql)
    'Set tmpReg.ActiveConnection = Nothing
    If (tmpReg.BOF Or tmpReg.EOF) Then
        GetAsiParam = ""
    Else
        sTip1 = "": sTip2 = "": sTip3 = "": sTip4 = "": sTip5 = "": sTip6 = ""
        With tmpReg
          Do While Not .EOF
            If !cNivOpe = "1" Then
                sTip1 = sTip1 & "'" & !cCodOpe & "',"
            ElseIf !cNivOpe = "2" Then
                sTip2 = sTip2 & "'" & !cCodOpe & "',"
            ElseIf !cNivOpe = "3" Then
                sTip3 = sTip3 & "'" & !cCodOpe & "',"
            ElseIf !cNivOpe = "4" Then
                sTip4 = sTip4 & "'" & !cCodOpe & "',"
            ElseIf !cNivOpe = "5" Then
                sTip5 = sTip5 & "'" & !cCodOpe & "',"
            ElseIf !cNivOpe = "6" Then
                sTip6 = sTip6 & "'" & !cCodOpe & "',"
            End If
            .MoveNext
          Loop
        End With
    End If
    tmpReg.Close
    Set tmpReg = Nothing
    If Len(sTip1) > 0 Then GetAsiParam = GetAsiParam & " And substring(" & vCampo & ",1,1) Not In (" & Left(sTip1, Len(sTip1) - 1) & ") "
    If Len(sTip2) > 0 Then GetAsiParam = GetAsiParam & " And substring(" & vCampo & ",1,2) Not In (" & Left(sTip2, Len(sTip2) - 1) & ") "
    If Len(sTip3) > 0 Then GetAsiParam = GetAsiParam & " And substring(" & vCampo & ",1,3) Not In (" & Left(sTip3, Len(sTip3) - 1) & ") "
    If Len(sTip4) > 0 Then GetAsiParam = GetAsiParam & " And substring(" & vCampo & ",1,4) Not In (" & Left(sTip4, Len(sTip4) - 1) & ") "
    If Len(sTip5) > 0 Then GetAsiParam = GetAsiParam & " And substring(" & vCampo & ",1,5) Not In (" & Left(sTip5, Len(sTip5) - 1) & ") "
    If Len(sTip6) > 0 Then GetAsiParam = GetAsiParam & " And " & vCampo & " Not In (" & Left(sTip6, Len(sTip6) - 1) & ") "
End Function

Public Function ClienteTipoCTS(ByVal pCuenta As String, oCon As ADODB.Connection) As String
    Dim tmpReg As ADODB.Recordset
    Dim tmpSql As String
    Set tmpReg = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    'oCon.AbreConexion
    
    tmpReg.CursorLocation = adUseClient
    tmpSql = "SELECT cCodInst FROM CaptacCTS WHERE cCtaCod = '" & pCuenta & "'"
    'ALPA 20080611
    Set tmpReg = oCon.Execute(tmpSql)
    'Set tmpReg = oCon.CargaRecordSet(tmpSql)
    'Set tmpReg.ActiveConnection = Nothing
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

Public Function ClienteTipoPersCol(ByVal nMovNro As Long, oCon As ADODB.Connection) As String
    Dim RegPer As ADODB.Recordset
    Dim tmpSql As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    'oCon.AbreConexion
    
    tmpSql = "SELECT C.nPersoneria FROM Captaciones C JOIN MovCap MC JOIN Mov M ON " _
        & "MC.nMovNro = M.nMovNro ON C.cCtaCod = MC.cCtaCod WHERE M.nMovNro = " & nMovNro
    
    Set RegPer = New ADODB.Recordset
    'RegPer.CursorLocation = adUseClient
    'Set RegPer = oCon.CargaRecordSet(tmpSql)
    'RegPer.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
    '**ALPA***20080609
    Set RegPer = oCon.Execute(tmpSql)
    'Set RegPer.ActiveConnection = Nothing
    If (RegPer.BOF Or RegPer.EOF) Then
        ClienteTipoPersCol = ""
    Else
        ClienteTipoPersCol = Trim(RegPer!nPersoneria)
    End If
    RegPer.Close
    Set RegPer = Nothing
End Function

Public Function ClienteTipoPers(ByVal pCuenta As String, oCon As ADODB.Connection) As String
    Dim RegPer As ADODB.Recordset
    Dim tmpSql As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    'oCon.AbreConexion
    
    Select Case Mid(pCuenta, 6, 3)
        Case Producto.gCapAhorros, Producto.gCapCTS, Producto.gCapPlazoFijo
            tmpSql = "SELECT nPersoneria FROM Captaciones WHERE cCtaCod = '" & pCuenta & "'"
        Case Else
            tmpSql = " SELECT Max(nPersPersoneria) nPersoneria FROM Persona AS P, ProductoPersona AS PC " & _
                     " WHERE P.cPersCod = PC.cPersCOd AND PC.cCtaCod = '" & pCuenta & "'" & _
                     " AND PC.nPrdPersRelac = " & ColocRelacPers.gColRelPersTitular
    End Select
    
    Set RegPer = New ADODB.Recordset
    'RegPer.CursorLocation = adUseClient
    'Set RegPer = oCon.CargaRecordSet(tmpSql)
    'RegPer.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
    '**ALPA***20080906
    Set RegPer = oCon.Execute(tmpSql)
    'Set RegPer.ActiveConnection = Nothing
    If (RegPer.BOF Or RegPer.EOF) Then
        ClienteTipoPers = ""
    Else
        ClienteTipoPers = Trim(RegPer!nPersoneria)
    End If
    RegPer.Close
    Set RegPer = Nothing
End Function

Public Function ExisAsiento(ByVal pCodContable As String, ByVal pFecha As Date, Optional pTipo As String = "0", Optional ByVal pNuePlan As Boolean = False, Optional oCon As ADODB.Connection) As Boolean
    Dim tmpReg As ADODB.Recordset
    Set tmpReg = New ADODB.Recordset
    Dim tmpSql As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    If pNuePlan Then
        tmpSql = "SELECT cCtaCnt FROM AsientoDN  " & _
            " WHERE cCtaCnt = '" & pCodContable & "' AND cTipo = '" & pTipo & "' AND datediff(dd, dFecha ,'" & Format(pFecha, "mm/dd/yyyy") & "') = 0 "
    Else
        tmpSql = "SELECT cCtaCnt FROM AsientoD  " & _
            " WHERE cCtaCnt = '" & pCodContable & "' AND cTipo = '" & pTipo & "' AND datediff(dd, dFecha ,'" & Format(pFecha, "mm/dd/yyyy") & "') = 0 "
    End If
    
    'Set tmpReg = oCon.CargaRecordSet(tmpSql)
    'ALPA 20080611
    Set tmpReg = oCon.Execute(tmpSql)
    If (tmpReg.BOF Or tmpReg.EOF) Then
        ExisAsiento = False
    Else
        ExisAsiento = True
    End If
    tmpReg.Close
    Set tmpReg = Nothing
End Function

Public Function ExisTipPer(ByVal pCodOpe As String, pConceptoCod As Long, Optional ByVal pNuePlan As Boolean = False, Optional oCon As ADODB.Connection) As Boolean
''''    Dim tmpReg As New ADODB.Recordset
''''    Dim tmpSql As String
''''    'Dim oCon As DConecta
''''    'Set oCon = New DConecta
''''
''''    'oCon.AbreConexion
''''
''''    tmpSql = "SELECT nPersoneria FROM OpeCtaNeg WHERE cOpeCod = '" & pCodOpe & "' And nConcepto = " & pConceptoCod
''''    tmpReg.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
''''    If (tmpReg.BOF Or tmpReg.EOF) Then
''''        ExisTipPer = False
''''    Else
''''        ExisTipPer = True
''''        Do While Not tmpReg.EOF
''''            If Trim(tmpReg!nPersoneria) = "0" Then
''''                ExisTipPer = False
''''                Exit Do
''''            End If
''''            tmpReg.MoveNext
''''        Loop
''''    End If
''''    tmpReg.Close
''''    Set tmpReg = Nothing
''''    'oCon.CierraConexion
''''    'Set oCon = Nothing
'***ALPA*****************************************20080604*********************
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    Set tmpReg = New ADODB.Recordset
    tmpSql = "SELECT nPersoneria FROM OpeCtaNeg WHERE cOpeCod = '" & pCodOpe & "' And nConcepto = " & pConceptoCod
    'ALPA***ASIENTO
    'tmpReg.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
    Set tmpReg = oCon.Execute(tmpSql)
    'Set tmpReg = oCon.CargaRecordSet(tmpSql)
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
    '***ALPA*****************************************20080604*********************
End Function

Public Function CuentaNombre(ByVal pCodCta As String, Optional ByVal pNuePlan As Boolean = False, Optional oCon As ADODB.Connection) As String
'''    Dim tmpReg As New ADODB.Recordset
'''    Dim tmpSql As String
'''    Dim pcodcta2 As String, pcodcta3 As String
'''    'Dim oCon As DConecta
'''    'Set oCon = New DConecta
'''    pCodCta = Trim(pCodCta)
'''    If Len(pCodCta) > 4 Then pcodcta2 = Left(pCodCta, Len(pCodCta) - 2)
'''    If Len(pCodCta) > 6 Then pcodcta3 = Left(pCodCta, Len(pCodCta) - 4)
'''
'''    'oCon.AbreConexion
'''
'''    tmpSql = "SELECT cCtaContDesc FROM  CtaCont WHERE cCtaContCod IN ('" & Left(pCodCta, 4) & "','" & _
'''        pCodCta & "','" & pcodcta2 & "','" & pcodcta3 & "') order by cCtaContCod"
'''    tmpReg.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
'''
'''    'Set tmpReg = oCon.CargaRecordSet(tmpSql)
'''    'tmpReg.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
'''
'''    If (tmpReg.BOF Or tmpReg.EOF) Then
'''        CuentaNombre = ""
'''    Else
'''        Do While Not tmpReg.EOF
'''            CuentaNombre = Trim(CuentaNombre) & " " & Trim(tmpReg!cCtaContDesc)
'''            tmpReg.MoveNext
'''        Loop
'''    End If
'''    tmpReg.Close
'''    Set tmpReg = Nothing
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    Dim pcodcta2 As String, pcodcta3 As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    pCodCta = Trim(pCodCta)
    If Len(pCodCta) > 4 Then pcodcta2 = Left(pCodCta, Len(pCodCta) - 2)
    If Len(pCodCta) > 6 Then pcodcta3 = Left(pCodCta, Len(pCodCta) - 4)

    'oCon.AbreConexion
    
    tmpSql = "SELECT cCtaContDesc FROM  CtaCont WHERE cCtaContCod IN ('" & Left(pCodCta, 4) & "','" & _
        pCodCta & "','" & pcodcta2 & "','" & pcodcta3 & "') order by cCtaContCod"
    '**ALPA****20080604
    'tmpReg.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
    Set tmpReg = oCon.Execute(tmpSql)
    'Set tmpReg = oCon.CargaRecordSet(tmpSql)
    'tmpReg.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText

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

Public Function CreditoPlazo(ByVal pCuenta As String, ByVal pCodLinC As String, ByVal pCodLinCJ As String) As String
    If Left(pCuenta, 2) = Right(gsCodAge, 2) Then  ' Para Judicial
       CreditoPlazo = Mid(Trim(pCodLinC), 5, 1)
    Else
       CreditoPlazo = Mid(Trim(pCodLinCJ), 5, 1)
    End If
End Function

Public Function CreditoFondo(ByVal pCuenta As String, ByVal pCodLinC As String, ByVal pCodLinCJ As String) As String
    If Left(pCuenta, 2) = Right(gsCodAge, 2) Then  '** Para Judicial
        CreditoFondo = Mid(Trim(pCodLinC), 6, 1)
    Else
        CreditoFondo = Mid(Trim(pCodLinCJ), 6, 1)
    End If
    If CreditoFondo = "1" Then                              'Recursos Propios
        CreditoFondo = "01"
    ElseIf CreditoFondo = "4" Then                          'BID
        CreditoFondo = "03"
    ElseIf CreditoFondo = "2" Or CreditoFondo = "3" Or CreditoFondo = "5" Or _
    CreditoFondo = "6" Then    ' Otros
        CreditoFondo = "02"
    Else
        CreditoFondo = ""
    End If
End Function

Public Function CreditoRecurso(ByVal pCuenta As String, ByVal pCodLinC As String, ByVal pCodLinCJ As String) As String
    Dim vTipoCred As String * 1
    Dim vTipoFond As String * 1
    Dim vVersion As String * 2
    Dim vOtroTipo As String * 2
    Dim vCodLinea As String
    
    If Left(pCuenta, 2) = Right(gsCodAge, 2) Then  ' Para Judicial
       vCodLinea = pCodLinC
    Else
       vCodLinea = pCodLinCJ
    End If
    
    'RegCta.Open tmpSql, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
    'If (RegCta.BOF Or RegCta.EOF) Then
    '    CreditoRecurso = ""
    'Else
        vTipoCred = Left(Trim(vCodLinea), 1)       'Mes
        vTipoFond = Mid(Trim(vCodLinea), 6, 1)     'Recursos Propios
        vVersion = Right(Trim(vCodLinea), 2)       'Version 01
        vOtroTipo = Mid(Trim(vCodLinea), 2, 2)     'Susy
        If Mid(vCodLinea, 3, 1) = "2" And _
          (Left(vCodLinea, 1) = "1" Or Left(vCodLinea, 1) = "2") Then 'Agricola
            CreditoRecurso = "04"
        'Comercial - RRPP
        'ElseIf vTipoCred = "1" And vTipoFond = "1" And vVersion = "01" Then
        '    CreditoRecurso = "05"
        ElseIf vTipoCred = "1" And vTipoFond = "1" Then
             CreditoRecurso = "05"
        ElseIf vTipoCred = "1" And vTipoFond <> "1" Then ' Parche para Foncodes
             CreditoRecurso = "02"
            
        'Mes - RRPP
        ElseIf vTipoCred = "2" And vTipoFond = "1" Then 'NO hay versiones 2000/08/31 Susy
            CreditoRecurso = "01"
        
        'Mes - Foncodes
        ElseIf vTipoCred = "2" And vTipoFond = "2" Then
            CreditoRecurso = "01"
        'Mes - Cofide , ya no hay versión 11/07/2000 - Susy
        ElseIf vTipoCred = "2" And vTipoFond = "5" Then
            CreditoRecurso = "03"
        ElseIf vTipoCred = "2" And vTipoFond = "3" Then
            CreditoRecurso = "02"
        ElseIf vTipoCred = "2" And vTipoFond = "6" Then
            CreditoRecurso = "06"
        'Mes - Bid
        ElseIf vTipoCred = "2" And vTipoFond = "4" Then
            CreditoRecurso = "01"
        'Consumo - RRPP   ' CAMBIO para LINEA CREDITO
        ElseIf vTipoCred = "3" And (vTipoFond = "1" Or vTipoFond = "3") And _
        (vOtroTipo = "01" Or vOtroTipo = "04") Then
            CreditoRecurso = "11"
        ElseIf vTipoCred = "3" And (vTipoFond = "1" Or vTipoFond = "3") And _
        vOtroTipo = "02" Then
            CreditoRecurso = "12"
        ElseIf vTipoCred = "3" And (vTipoFond = "1" Or vTipoFond = "3") And _
        vOtroTipo = "03" Then
            CreditoRecurso = "13"
        '***** Parche Administrativos - Arturo
        ElseIf vTipoCred = "3" And (vTipoFond = "1" Or vTipoFond = "3") And _
        vOtroTipo = "20" Then
            CreditoRecurso = "20"
        '***** Parche Administrativos
        Else
            CreditoRecurso = ""
        End If
        
        
    'End If
    'RegCta.Close
    'Set RegCta = Nothing
End Function

Public Function CreditoCJ(ByVal pCodAge As String, oCon As ADODB.Connection) As String
    Dim RegCta As ADODB.Recordset
    Set RegCta = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    Dim tmpSql As String
    
    'oCon.AbreConexion
    tmpSql = "SELECT nRanIniTab FROM TablaCod WHERE substring(ccodtab,1,2) = '47' AND cValor = '" & pCodAge & "'"
    'RegCta.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
    '**ALPA***20080906
    Set RegCta = oCon.Execute(tmpSql)
    
    If (RegCta.BOF Or RegCta.EOF) Then
        CreditoCJ = ""
    Else
        CreditoCJ = Trim(String(2 - Len(Round(RegCta!nRanIniTab, 0)), "0")) & Trim(Round(RegCta!nRanIniTab, 0))
    End If
    RegCta.Close
    Set RegCta = Nothing
End Function

Public Function CreditoRE(ByVal pCuenta As String, ByVal pCodLinC As String, ByVal pCodLinCJ As String) As String
    Dim vCodLinea As String
    
    vCodLinea = pCodLinC
    If Mid(Trim(vCodLinea), 6, 1) = "1" Then
        CreditoRE = "1"
    ElseIf Mid(Trim(vCodLinea), 6, 1) = "2" Or Mid(Trim(vCodLinea), 6, 1) = "3" Then
        CreditoRE = "2"
    ElseIf Mid(Trim(vCodLinea), 6, 1) = "4" Then
        CreditoRE = "3"
    Else
        CreditoRE = ""
    End If
End Function

Public Function AsientoParche(pAsiento As String, Optional ByVal pNuePlan As Boolean = False, Optional oCon As ADODB.Connection) As String
''''    Dim tmpRs As New ADODB.Recordset
''''    Dim tmpSql As String
''''    'Dim oCon As DConecta
''''    'Set oCon = New DConecta
''''
''''    'oCon.AbreConexion
''''
''''    If pNuePlan Then
''''        tmpSql = " SELECT cCodCntNue" & _
''''            " FROM AsientoPuenteN " & _
''''            " WHERE cCodCnt = '" & Trim(pAsiento) & "'"
''''    Else
''''        tmpSql = " SELECT cCodCntNue" & _
''''            " FROM AsientoPuente " & _
''''            " WHERE cCodCnt = '" & Trim(pAsiento) & "'"
''''    End If
''''    'Set tmpRs = oCon.CargaRecordSet(tmpSql)
''''    Set tmpRs = New ADODB.Recordset
''''    tmpRs.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
''''    If (tmpRs.BOF Or tmpRs.EOF) Then
''''        AsientoParche = ""
''''    Else
''''        AsientoParche = Trim(tmpRs!cCodCntNue)
''''    End If
''''    tmpRs.Close
''''    Set tmpRs = Nothing
'***ALPA***20080604**********************************************************************
Dim tmpRs As New ADODB.Recordset
    Dim tmpSql As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    'oCon.AbreConexion
    
    If pNuePlan Then
        tmpSql = " SELECT cCodCntNue" & _
            " FROM AsientoPuenteN " & _
            " WHERE cCodCnt = '" & Trim(pAsiento) & "'"
    Else
        tmpSql = " SELECT cCodCntNue" & _
            " FROM AsientoPuente " & _
            " WHERE cCodCnt = '" & Trim(pAsiento) & "'"
    End If
    'Set tmpRs = oCon.CargaRecordSet(tmpSql)
    Set tmpRs = New ADODB.Recordset
    'ALPA***ASIENTO
    'tmpRs.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
    Set tmpRs = oCon.Execute(tmpSql)
    If (tmpRs.BOF Or tmpRs.EOF) Then
        AsientoParche = ""
    Else
        AsientoParche = Trim(tmpRs!cCodCntNue)
    End If
    tmpRs.Close
    Set tmpRs = Nothing
    '***ALPA***20080604**********************************************************************
End Function

Public Function ExisCtaCnt(ByVal pCtaCnt As String, Optional ByVal pNuePlan As Boolean = False, Optional oCon As ADODB.Connection) As Boolean
'''    Dim tmpReg As New ADODB.Recordset
'''    Dim tmpSql As String
'''    'Dim oCon As DConecta
'''    'Set oCon = New DConecta
'''
'''    'oCon.AbreConexion
'''    tmpSql = "SELECT cCtaContDesc FROM CtaCont WHERE cCtaContCod LIKE '" & pCtaCnt & "%' and nCtaEstado = 1"
'''    'Set tmpReg = oCon.CargaRecordSet(tmpSql)
'''    tmpReg.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
'''    If (tmpReg.BOF And tmpReg.EOF) Then
'''        ExisCtaCnt = False
'''    Else
'''        ExisCtaCnt = IIf(tmpReg.RecordCount = 1, True, False)
'''    End If
'''    tmpReg.Close
'''    Set tmpReg = Nothing
'''    'oCon.CierraConexion
'''    'Set oCon = Nothing
''' ALPA***20080604****************************************************************************
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    'oCon.AbreConexion
    tmpSql = "SELECT cCtaContDesc FROM CtaCont WHERE cCtaContCod LIKE '" & pCtaCnt & "%' and nCtaEstado = 1"
    'Set tmpReg = oCon.CargaRecordSet(tmpSql)
    'ALPA***ASIENTO
    'tmpReg.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
    Set tmpReg = oCon.Execute(tmpSql)
    If (tmpReg.BOF And tmpReg.EOF) Then
        ExisCtaCnt = False
    Else
        ExisCtaCnt = IIf(tmpReg.RecordCount = 1, True, False)
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

Public Function VarCJ(ByVal pCmac As String, ByVal pCodCta As String, oCon As ADODB.Connection) As String
    Dim RegCta As ADODB.Recordset
    Set RegCta = New ADODB.Recordset
    Dim tmpSql As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta

    'oCon.AbreConexion
    
    tmpSql = " SELECT cCodCnt AS Campo FROM CuentaIFEspecial " & _
             " WHERE cCtaCod = '" & pCodCta & "' "
    'Set RegCta = oCon.CargaRecordSet(tmpSql)
    'RegCta.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
    '**ALPA***20080906
    Set RegCta = oCon.Execute(tmpSql)
    If (RegCta.BOF Or RegCta.EOF) Then
        VarCJ = ""
    Else
        VarCJ = Trim(RegCta!Campo)
    End If
    RegCta.Close
    Set RegCta = Nothing
End Function

Public Function VarCR(ByVal sMoneda As String, ByVal sCodCta As String, oCon As ADODB.Connection) As String
    Dim RegCta As ADODB.Recordset
    Dim tmpSql As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    'oCon.AbreConexion
    
    tmpSql = "SELECT cValor2 AS Campo FROM CuentaIFEspecial " & _
        " WHERE cNumCta = '" & sCodCta & "' And cCodCtaEsp LIKE '0103%'"
    'RegCta.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
    '**ALPA***20080906
    Set RegCta = oCon.Execute(tmpSql)
    'Set RegCta = New ADODB.Recordset
    'RegCta.CursorLocation = adUseClient
    'Set RegCta = oCon.CargaRecordSet(tmpSql)
    'Set RegCta.ActiveConnection = Nothing
    If (RegCta.BOF Or RegCta.EOF) Then
        VarCR = ""
    Else
        VarCR = Trim(RegCta!Campo)
    End If
    RegCta.Close
    Set RegCta = Nothing
End Function

Public Function VarInstitucionFinanciera(ByVal sCodCta As String, ByVal nTipoIF As CGTipoIF, oCon As ADODB.Connection) As String
    Dim RegCta As ADODB.Recordset
    Dim tmpSql As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    'oCon.AbreConexion
    
    
    '*** PEAC 20111107
    'tmpSql = "SELECT I.cSubCtaContCod AS Campo FROM InstitucionFinanc I JOIN ProductoPersona PP " _
        & "JOIN Producto P ON PP.cCtaCod = P.cCtaCod ON I.cPersCod = PP.cPersCod WHERE " _
        & "P.cCtaCod = '" & sCodCta & "' And PP.nPrdPersRelac = " & gCapRelPersTitular & " And " _
        & "I.cIFTpo = '" & Format$(nTipoIF, "00") & "'"
        
    tmpSql = "exec stp_sel_VarInstitucionFinanciera '" & sCodCta & "','" & Format$(nTipoIF, "00") & "'"
        
    '*** FIN PEAC
        
    Set RegCta = New ADODB.Recordset
    'RegCta.CursorLocation = adUseClient
    'Set RegCta = oCon.CargaRecordSet(tmpSql)
    'Set RegCta.ActiveConnection = Nothing
    'RegCta.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
    '**ALPA***20080906
    Set RegCta = oCon.Execute(tmpSql)
    If (RegCta.BOF Or RegCta.EOF) Then
        VarInstitucionFinanciera = ""
    Else
        VarInstitucionFinanciera = Trim(RegCta!Campo)
    End If
    RegCta.Close
    Set RegCta = Nothing
End Function

Public Function VarInstFinanMov(ByVal pnMovNro As Long, ByVal nTipoIF As CGTipoIF, oCon As ADODB.Connection) As String
    Dim RegCta As ADODB.Recordset
    Dim tmpSql As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    'oCon.AbreConexion
    
    tmpSql = " Select   i.cSubCtaContCod  as Campo " _
        & "    from     movcmac mc   " _
        & "             join InstitucionFinanc i on i.cperscod=mc.cPerscod   " _
        & "             join persona p on p.cperscod = i.cperscod " _
        & "     where nmovnro =" & pnMovNro & " and I.cIFTpo = '" & Format$(nTipoIF, "00") & "'"
        
    Set RegCta = New ADODB.Recordset
    'RegCta.CursorLocation = adUseClient
    'Set RegCta = oCon.CargaRecordSet(tmpSql)
    'Set RegCta.ActiveConnection = Nothing
    'RegCta.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
    '**ALPA***20080906
    Set RegCta = oCon.Execute(tmpSql)
    If (RegCta.BOF Or RegCta.EOF) Then
        VarInstFinanMov = ""
    Else
        VarInstFinanMov = Trim(RegCta!Campo)
    End If
    RegCta.Close
    Set RegCta = Nothing
End Function

Public Function VarBC(ByVal nMovNro As Long, ByVal sCtaConCod As String, oCon As ADODB.Connection) As String
Dim RegCta As ADODB.Recordset
Dim tmpSql As String, sCtaCod As String
'Dim oCon As DConecta
'Set oCon = New DConecta
'oCon.AbreConexion

sCtaCod = Mid(sCtaConCod, 1, InStr(1, sCtaConCod, "BC", vbTextCompare) - 1)

tmpSql = "SELECT ISNULL(cCtaIFSubCta,'') Campo FROM Mov M JOIN MovOpeVarias V JOIN CtaIFFiltro F ON " _
    & "SUBSTRING(V.cReferencia,4,13) = F.cPersCod AND LEFT(V.cReferencia,2) = F.cIFTpo " _
    & "AND SUBSTRING(V.cReferencia,18,7) = F.cCtaIfCod ON M.nMovNro = V.nMovNro " _
    & "WHERE M.nMovNro = " & nMovNro
'Set RegCta = oCon.CargaRecordSet(tmpSql)
'RegCta.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText
    '**ALPA***20080906
    Set RegCta = oCon.Execute(tmpSql)

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

Public Function VarFF(ByVal pCuenta As String, ByVal pCodLinC As String, ByVal pCodLinCJ As String, oCon As ADODB.Connection) As String
    Dim RegCta As New ADODB.Recordset
    Dim tmpSql1 As String, sServConsol As String
    'Dim oCon As DConecta
    
    'ALPA 20090206***************************************************************
    'Dim loConstSist As NConstSistemas
    'Set loConstSist = New NConstSistemas
    sServConsol = "DBConsolidada.." 'oConstSist.LeeConstSistema(gConstSistServCentralRiesgos)
    'Set loConstSist = Nothing
    '****************************************************************************

    'Set oCon = New DConecta
    'oCon.AbreConexion
    '** EQUIVALENTE FUENTE FINANCIAMIENTO
''    tmpSql1 = "Select cCtaCont FROM " & sServConsol & "ColocLineaCreditoEquiv WHERE cLineaCred = '" & Mid(pCodLinC, 1, 4) & "' "
''    'Set RegCta = oCon.CargaRecordSet(tmpSql1)
''    RegCta.Open tmpSql1, oCon, adOpenStatic, adLockReadOnly, adCmdText
''    If (RegCta.BOF Or RegCta.EOF) Then
''        VarFF = ""
''    Else
''       VarFF = RegCta!cCtaCont
''    End If
    tmpSql1 = "Select cCtaCont FROM " & sServConsol & "ColocLineaCreditoEquiv WHERE cLineaCred = '" & Mid(pCodLinC, 1, 4) & "' "
    'ALPA 20080604
    Set RegCta = oCon.Execute(tmpSql1)
    If (RegCta.BOF Or RegCta.EOF) Then
        VarFF = ""
    Else
       VarFF = RegCta!cCtaCont
    End If
End Function

Public Function VarPD(ByVal pCuenta As String, ByVal pCodLinC As String, ByVal pCodLinCJ As String) As String
    Dim vCodLinea As String
    
    vCodLinea = pCodLinC
        
    vCad = Mid(Trim(pCuenta), 6, 3)
    If (vCad = "101" Or vCad = "201") And Mid(Trim(vCodLinea), 6, 1) = "2" Then
        VarPD = "05"
    ElseIf vCad = "101" Or vCad = "201" Then
        VarPD = "01"
    ElseIf vCad = "301" Or vCad = "302" Or vCad = "303" Or vCad = "304" Or vCad = "320" Then
        VarPD = "02"
    ElseIf vCad = "102" Or vCad = "202" Then
        VarPD = "03"
    ElseIf vCad = "401" Or vCad = "423" Or vCad = "402" Then
        VarPD = "02"
    Else
        VarPD = ""
    End If
End Function

Public Function VarAsientoProdEquiv(ByVal psTpoCredCod As String, ByVal psTpoProdCod As String, ByVal pnTpoInstCorp, ByVal psEquivTipo As String, oCon As ADODB.Connection) As String
Dim RegCta As New ADODB.Recordset
'Dim tmpSql As String
Dim lsSQL As String
  
    '**Modificado por DAOR 20100623 ***********************************
    'tmpSql = "SELECT cEquivAsiento AS Campo FROM AsientoProdEquiv " & _
    '            "WHERE cProducto ='" & pProducto & "' AND cEquivTpo = '" & pEquivTipo & "' "
    lsSQL = "exec B2_stp_sel_EquivalenciasParaAsiento '" & psTpoCredCod & "','" & psTpoProdCod & "'," & pnTpoInstCorp
    '******************************************************************
    
    If (sTpoCredCod <> psTpoCredCod Or sTpoProdCod <> psTpoProdCod Or nTpoInstCorp <> pnTpoInstCorp) Then
        sTpoCredCod = psTpoCredCod
        sTpoProdCod = psTpoProdCod
        nTpoInstCorp = pnTpoInstCorp
        Set RegCta = oCon.Execute(lsSQL)
        If (RegCta.BOF Or RegCta.EOF) Then
            sTC = "":   sCD = "": sIK = "": sSC = ""
        Else
            sTC = IIf(IsNull(RegCta!cEquivTC), "", Trim(RegCta!cEquivTC))
            sCD = IIf(IsNull(RegCta!cEquivCD), "", Trim(RegCta!cEquivCD))
            sIK = IIf(IsNull(RegCta!cEquivIK), "", Trim(RegCta!cEquivIK))
            sSC = IIf(IsNull(RegCta!cEquivSC), "", Trim(RegCta!cEquivSC))
            sTL = IIf(IsNull(RegCta!cEquivTL), "", Trim(RegCta!cEquivTL))
        End If
        RegCta.Close
        Set RegCta = Nothing
    End If
    
    Select Case psEquivTipo
        Case "TC"
            VarAsientoProdEquiv = sTC
        Case "CD"
            VarAsientoProdEquiv = sCD
        Case "IK"
            VarAsientoProdEquiv = sIK
        Case "SC"
            VarAsientoProdEquiv = sSC
        Case "TL"
            VarAsientoProdEquiv = sTL
    End Select
    
End Function
Public Sub ObtenerTipoCambioLeasing(ByVal psCtaCod As String, oCon As ADODB.Connection, ByRef lnTipoCambioCompraL As Currency, ByRef lnTipoCambioVentaL As Currency, ByRef lnTipoCambioFijoL As Currency)
        Dim sSqlLeasing As String
        Dim sSqlTCL As String
        Dim lobjRS As ADODB.Recordset
        Dim lobjRSTC As ADODB.Recordset
        
        Set lobjRS = New ADODB.Recordset
        Set lobjRSTC = New ADODB.Recordset
                            
                            
        lnTipoCambioCompraL = 0
        lnTipoCambioVentaL = 0
        lnTipoCambioFijoL = 0

                                                
        sSqlLeasing = "select top 1 dFecha=convert(DateTime,left(cMovNro,8)) "
        sSqlLeasing = sSqlLeasing & " from DBCmacMaynas..Mov M"
        sSqlLeasing = sSqlLeasing & " inner join DBCmacMaynas..MovRef MR on MR.nMovNro=M.nMovNro"
        sSqlLeasing = sSqlLeasing & " inner join SAF..OC_Core_nmovnro OC on OC.num_core_nmovnro=MR.nMovNroRef"
        sSqlLeasing = sSqlLeasing & " Inner Join"
        sSqlLeasing = sSqlLeasing & " ("
        sSqlLeasing = sSqlLeasing & " select Saf_numero_operacion_saf, Saf_numero_operacion_core"
        sSqlLeasing = sSqlLeasing & " from SAF..IntSaf_Activacion_Cronograma"
        sSqlLeasing = sSqlLeasing & " group by Saf_numero_operacion_saf, Saf_numero_operacion_core"
        sSqlLeasing = sSqlLeasing & " ) OI on OI.Saf_numero_operacion_saf=num_operacion"
        sSqlLeasing = sSqlLeasing & " where M.cOpeCod='562502' and M.nMovFlag=0 "
        sSqlLeasing = sSqlLeasing & " and Saf_numero_operacion_core='" & psCtaCod & "'"
        Set lobjRS = oCon.Execute(sSqlLeasing)
        If Not RSVacio(lobjRS) Then
            sSqlTCL = " select nValFijo,nValVent,nValComp "
            sSqlTCL = sSqlTCL & " From TipoCambio "
            sSqlTCL = sSqlTCL & "Where datediff(d,dFecCamb,'" & Format(lobjRS!dFecha, "YYYY/MM/DD") & "')=0 "
            Set lobjRSTC = oCon.Execute(sSqlTCL)
            If Not RSVacio(lobjRSTC) Then
                lnTipoCambioCompraL = lobjRSTC!nValComp
                lnTipoCambioVentaL = lobjRSTC!nValVent
                lnTipoCambioFijoL = lobjRSTC!nValFijo
            End If
        End If
        Set lobjRS = Nothing
        Set lobjRSTC = Nothing
End Sub
Public Sub IngresarAsientoDN(pdHoraGrab As String, psCodConta As String, pnMontoOperacionD As Double, pnMontoOperacionH As Double, psTipo As String, psCodAge As String, Optional ByVal pnMovNro As Long = -1, Optional psOpecod As String = "", Optional psCtaCod As String = "", Optional oCon As ADODB.Connection = Nothing, Optional ByVal pnTipoCambio As Currency = 0#)
Dim ssql As String
Dim oRs1 As ADODB.Recordset
Set oRs1 = New ADODB.Recordset
Dim psCodContaTemporal2107 As String
Dim psCodContaTemporal2102 As String
Dim psFecha As String
Dim oRs2 As ADODB.Recordset
Set oRs2 = New ADODB.Recordset
Dim lbInactivas As Boolean
Dim bEsMovCCE As Boolean 'PASI20161202 CCE
Dim bEsOpeCodChe As Boolean 'VAPA20170202
Dim bOpeCodValido As Boolean 'VAPA20170204

'             If psCtaCod = "109012321000034771" Or psCtaCod = "109012321000426075" Then
'             MsgBox psCtaCod
'             End If
   'If Mid(psCtaCod, 6, 3) = "232" Or Mid(psCtaCod, 6, 3) = "233" Or Mid(psCtaCod, 6, 3) = "234" Then
   
   bEsOpeCodChe = EsOpeCodCheCCE(psOpecod, oCon) 'VAPA20170202
   bEsMovCCE = EsMovDeChequeCCE(pnMovNro, oCon) 'PASI20161202 CCE
   If bEsMovCCE = True And bEsOpeCodChe = True Then 'VAPA20170203
                    'BARRABASA  pnMontoOperacionD pnMovNro psOpecod
                    
                    bOpeCodValido = EsOpeCodCambioACCE(pnMovNro, psOpecod, pnMontoOperacionD, oCon)
                    
   End If
   
   
   'JUEZ 20150331 Para operaciones de seguro de tarjetas redondear a un decimal ******************************
   If psTipo = "3" And (psOpecod = gAhoCargoAfilSegTarjeta Or psOpecod = gCTSCargoAfilSegTarjeta Or _
                        psOpecod = 200375 Or psOpecod = 220311 Or psOpecod = 300437 Or _
                        psOpecod = gAhoDepAfilSegTarjeta Or psOpecod = gCTSDepAfilSegTarjeta) Then
        pnMontoOperacionD = Round(pnMontoOperacionD, 1)
        pnMontoOperacionH = Round(pnMontoOperacionH, 1)
   End If
   'END JUEZ *************************************************************************************************
   
   lbInactivas = 0
   If Mid(psCtaCod, 6, 3) = "232" Then
        ssql = "select * from CaptacInactivasTotal where substring(cCtaCod,6,3)='232' and cCtaCod='" & psCtaCod & "' and Convert(Date,dFecha)=DateADD(day,-day(Convert(Date,'" & pdHoraGrab & "')),Convert(Date,'" & pdHoraGrab & "')) "
        Set oRs1 = oCon.Execute(ssql)
        If Not (oRs1.BOF Or oRs1.EOF) Then
             lbInactivas = 1
             psFecha = Mid(pdHoraGrab, 7, 4) & Mid(pdHoraGrab, 1, 2) & Mid(pdHoraGrab, 4, 2) '12/30/2014

             ssql = "Select C.nPersoneria,MC.cCtaCod,CS.nSaldoCnt,CS.nInteres,binactiva=isnull(CS.nInmovilizada,0),cCodInst=isnull(cCodInst,0) "
             ssql = ssql & "From Mov M  inner join MovCap MC on M.nMovNro=MC.nMovNro and M.cOpeCod=MC.cOpeCod"
             ssql = ssql & "            inner join CaptacInactivasTotal CS on MC.cCtaCod=CS.cCtaCod "
             ssql = ssql & "                and Convert(Date,CS.dFecha)=DateADD(day,-day(Convert(Date,'" & pdHoraGrab & "')),Convert(Date,'" & pdHoraGrab & "'))"
             ssql = ssql & "            inner join Captaciones C on C.cCtaCod=CS.cCtaCod "
             ssql = ssql & "            left join captaccts cts on c.cctacod = cts.cctacod "
             ssql = ssql & "Where MC.cCtaCod='" & psCtaCod & "' and M.cOpeCod='" & psOpecod & "' and M.cOpeCod in (select cOpeCod from OpeTpoNegInactivas) "
             ssql = ssql & "    and cMovNro like left('" & psFecha & "',6)+'%' "
             ssql = ssql & "    and M.nMovNro = " & pnMovNro & " and MC.cCtaCod not in (select cCtaCod from AsientoDN where  Datediff(MONTH,dFecha,'" & psFecha & "')=0  and cCtaCod='" & psCtaCod & "' and cOpeCod in (select cOpeCod from OpeTpoNegInactivas)) "
             Set oRs2 = oCon.Execute(ssql)
             If Not (oRs2.BOF Or oRs2.EOF) Then
                    If psOpecod <> "200802" Then
                    '0-Moneda
                    '3-Dolares
                    psCodContaTemporal2107 = "21" & Mid(psCtaCod, 9, 1) & "701" & IIf(oRs2!nPersoneria = "1", "0101", IIf(oRs2!nPersoneria = "2", "0102", "02")) & psCodAge 'Right(psCodConta, 2)
                    psCodContaTemporal2102 = CreaCuentaAhorroDN(psCtaCod, pdHoraGrab, oRs2!nPersoneria, 0, oRs2!cCodInst, oCon)
                    
                    'Capital
                    If oRs1!nSaldoCnt > 0 Then
                        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) "
                        'ssql = ssql & " VALUES('" & pdHoraGrab & "','" & psCodContaTemporal2107 & "'," & oRs1!nSaldoCnt & ", " & 0 & ", '0' ,'" & psCodAge & "'," & pnMovNro & ",'" & psOpeCod & "','" & psCtaCod & "')"'/**Comentado PASI20161202 ***/
                        ssql = ssql & " VALUES('" & pdHoraGrab & "','" & IIf(bEsMovCCE And bOpeCodValido And Not psCodContaTemporal2107 = "41" & Mid(psCodContaTemporal2107, 3, 1) & "103050101" & psCodAge, "25" & Mid(psCodContaTemporal2107, 3, 1) & "4190502", psCodContaTemporal2107) & "'," & oRs1!nSaldoCnt & ", " & 0 & ", '0' ,'" & psCodAge & "'," & pnMovNro & ",'" & psOpecod & "','" & psCtaCod & "')" 'PASI20161202 CCE'VAPA20170202
                        oCon.Execute ssql
                        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) "
                        ssql = ssql & " VALUES('" & pdHoraGrab & "','" & psCodContaTemporal2102 & "'," & 0 & ", " & oRs1!nSaldoCnt & ", '0' ,'" & psCodAge & "'," & pnMovNro & ",'" & psOpecod & "','" & psCtaCod & "')"
                        oCon.Execute ssql
                    End If
                    
                    'Interes
                    If oRs1!nInteres > 0 Then
                        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) "
                        'ssql = ssql & " VALUES('" & pdHoraGrab & "','" & psCodContaTemporal2107 & "'," & oRs1!nInteres & ", " & 0 & ", '0' ,'" & psCodAge & "'," & pnMovNro & ",'" & psOpeCod & "','" & psCtaCod & "')"  '/**Comentado PASI20161202 ***/
                        ssql = ssql & " VALUES('" & pdHoraGrab & "','" & IIf(bEsMovCCE And bOpeCodValido And Not psCodContaTemporal2107 = "41" & Mid(psCodContaTemporal2107, 3, 1) & "103050101" & psCodAge, "25" & Mid(psCodContaTemporal2107, 3, 1) & "4190502", psCodContaTemporal2107) & "'," & oRs1!nInteres & ", " & 0 & ", '0' ,'" & psCodAge & "'," & pnMovNro & ",'" & psOpecod & "','" & psCtaCod & "')" 'PASI20161202 CCE 'VAPA20170202
                        oCon.Execute ssql
                        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) "
                        ssql = ssql & " VALUES('" & pdHoraGrab & "','" & psCodContaTemporal2102 & "'," & 0 & ", " & oRs1!nInteres & ", '0' ,'" & psCodAge & "'," & pnMovNro & ",'" & psOpecod & "','" & psCtaCod & "')"
                        oCon.Execute ssql
                    End If
                    If Mid(psCtaCod, 9, 1) = "2" Then
                        'Capital
                        If oRs1!nSaldoCnt > 0 Then
                            ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) "
                            'ssql = ssql & " VALUES('" & pdHoraGrab & "','" & psCodContaTemporal2107 & "'," & oRs1!nSaldoCnt * pnTipoCambio & ", " & 0 & ", '3' ,'" & psCodAge & "'," & pnMovNro & ",'" & psOpeCod & "','" & psCtaCod & "')" '/**Comentado PASI20161202 ***/
                            ssql = ssql & " VALUES('" & pdHoraGrab & "','" & IIf(bEsMovCCE And bOpeCodValido And Not psCodContaTemporal2107 = "41" & Mid(psCodContaTemporal2107, 3, 1) & "103050101" & psCodAge, "25" & Mid(psCodContaTemporal2107, 3, 1) & "4190502", psCodContaTemporal2107) & "'," & oRs1!nSaldoCnt * pnTipoCambio & ", " & 0 & ", '3' ,'" & psCodAge & "'," & pnMovNro & ",'" & psOpecod & "','" & psCtaCod & "')" 'PASI20161202 CCE'VAPA20170202
                            oCon.Execute ssql
                            ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) "
                            ssql = ssql & " VALUES('" & pdHoraGrab & "','" & psCodContaTemporal2102 & "'," & 0 & ", " & oRs1!nSaldoCnt * pnTipoCambio & ", '3' ,'" & psCodAge & "'," & pnMovNro & ",'" & psOpecod & "','" & psCtaCod & "')"
                            oCon.Execute ssql
                        End If
                        'Interes
                        If oRs1!nInteres > 0 Then
                            ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) "
                            'ssql = ssql & " VALUES('" & pdHoraGrab & "','" & psCodContaTemporal2107 & "'," & oRs1!nInteres * pnTipoCambio & ", " & 0 & ", '3' ,'" & psCodAge & "'," & pnMovNro & ",'" & psOpeCod & "','" & psCtaCod & "')" '/**Comentado PASI20161202 ***/
                            ssql = ssql & " VALUES('" & pdHoraGrab & "','" & IIf(bEsMovCCE And bOpeCodValido And Not psCodContaTemporal2107 = "41" & Mid(psCodContaTemporal2107, 3, 1) & "103050101" & psCodAge, "25" & Mid(psCodContaTemporal2107, 3, 1) & "4190502", psCodContaTemporal2107) & "'," & oRs1!nInteres * pnTipoCambio & ", " & 0 & ", '3' ,'" & psCodAge & "'," & pnMovNro & ",'" & psOpecod & "','" & psCtaCod & "')" 'PASI20161202 CCE'VAPA20170202
                            oCon.Execute ssql
                            ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) "
                            ssql = ssql & " VALUES('" & pdHoraGrab & "','" & psCodContaTemporal2102 & "'," & 0 & ", " & oRs1!nInteres * pnTipoCambio & ", '3' ,'" & psCodAge & "'," & pnMovNro & ",'" & psOpecod & "','" & psCtaCod & "')"
                            oCon.Execute ssql
                        End If
                    End If
                End If 'If psOpeCod <> "200802" Then
             End If
        Else
        If psOpecod = "200802" Then
            lbInactivas = 0
        End If
        End If
        Set oRs1 = Nothing
   End If
   
   If (lbInactivas = 0 And psOpecod = "200802") Or psOpecod <> "200802" Then
         ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge, nMovNro, cOpeCod, cCtaCod) "
         'ssql = ssql & " VALUES('" & pdHoraGrab & "','" & psCodConta & "'," & pnMontoOperacionD & ", " & pnMontoOperacionH & ", '" & psTipo & "' ,'" & psCodAge & "'" 'Modificado PASI20140430 '/***Comentado PASI20161202**/
         'sSql = sSql & " VALUES('" & pdHoraGrab & "','" & IIf(bEsMovCCE And Not pnMontoOperacionD = 0 And Not psCodConta = "41" & Mid(psCodConta, 3, 1) & "103050101" & Right(psCodConta, 2), "25" & Mid(psCodConta, 3, 1) & "4190502", psCodConta) & "'," & pnMontoOperacionD & ", " & pnMontoOperacionH & ", '" & psTipo & "' ,'" & psCodAge & "'"  'PASI20161202 CCE
         ssql = ssql & " VALUES('" & pdHoraGrab & "','" & IIf(bEsMovCCE And bOpeCodValido And Not pnMontoOperacionD = 0 And Not psCodConta = "41" & Mid(psCodConta, 3, 1) & "103050101" & Right(psCodConta, 2), "25" & Mid(psCodConta, 3, 1) & "4190502", psCodConta) & "'," & pnMontoOperacionD & ", " & pnMontoOperacionH & ", '" & psTipo & "' ,'" & psCodAge & "'"  'Agregado VAPA20170202
         
                
         
         
         'Agregado PASI20140430
         If pnMovNro = -1 Then
              ssql = ssql & ", Null "
         Else
              ssql = ssql & "," & pnMovNro
         End If
         If psOpecod = "" Then
              ssql = ssql & ", Null "
         Else
              ssql = ssql & ",'" & psOpecod & "'"
         End If
         If psCtaCod = "" Then
              ssql = ssql & ", Null "
         Else
              ssql = ssql & ",'" & psCtaCod & "'"
         End If
         ssql = ssql & ")"
         'END PASI
        oCon.Execute ssql
  End If
End Sub
Public Function CreaCuentaAhorro(ByVal psCtaCod As String, ByVal pdFecha As Date, ByVal pnPersoneria As Integer, ByVal pbInactiva As Integer, ByVal psCodInst As String, oCon As ADODB.Connection) As String
Dim ssql As String
Dim oRs1 As ADODB.Recordset
Dim psCtaContCod As String
Set oRs1 = New ADODB.Recordset

 ssql = ssql & " SELECT PC.CCTACOD"
 ssql = ssql & " FROM DBConsolidada..ProductoBloqueosConsol PC "
 ssql = ssql & " WHERE cMovNroDbl IS NULL AND nBlqMotivo = 3 And PC.cCtaCod Like '_____233%' "
 ssql = ssql & "        and Convert(Date,dCierre) = '" & Format(pdFecha, "YYYY/MM/DD") & "'"
 Set oRs1 = oCon.Execute(ssql)
 If (oRs1.BOF Or oRs1.EOF) Then
    If Mid(psCtaCod, 6, 3) = "232" Then
       If pbInactiva = 1 Then
          psCtaContCod = "21" & Mid(Trim(psCtaCod), 9, 1) & "202"
       Else
          psCtaContCod = "21" & Mid(Trim(psCtaCod), 9, 1) & "201"
       End If
    ElseIf Mid(psCtaCod, 6, 3) = "233" Then
          psCtaContCod = "21" & Mid(Trim(psCtaCod), 9, 1) & "303"
    ElseIf Mid(psCtaCod, 6, 3) = "234" Then
          psCtaContCod = "21" & Mid(Trim(psCtaCod), 9, 1) & "305"
    End If
    If (pnPersoneria = 1 Or pnPersoneria = 2) And Mid(psCtaCod, 6, 3) <> "234" Then
        psCtaContCod = psCtaContCod & "01"
    End If
    If (pnPersoneria = 2 Or pnPersoneria = 3) Then
        psCtaContCod = psCtaContCod & "02"
    Else
        If Mid(psCtaCod, 6, 3) <> "234" Then
            psCtaContCod = psCtaContCod & "01"
        Else
            If psCodInst = "1090100012521" Then
                psCtaContCod = psCtaContCod & "01"
            Else
                psCtaContCod = psCtaContCod & "02"
            End If
        End If
    End If
    psCtaContCod = psCtaContCod & Mid(Trim(psCtaCod), 4, 2)
 Else
    psCtaContCod = "21" & Mid(Trim(psCtaCod), 9, 1) & "70401"
    psCtaContCod = psCtaContCod & "01"
    psCtaContCod = psCtaContCod & Mid(Trim(psCtaCod), 4, 2)
 End If
End Function
Public Function CreaCuentaAhorroDN(ByVal psCtaCod As String, ByVal pdFecha As Date, ByVal pnPersoneria As Integer, ByVal pbInactiva As Integer, ByVal psCodInst As String, oCon As ADODB.Connection) As String
Dim ssql As String
Dim oRs1 As ADODB.Recordset
Dim psCtaContCod As String
Set oRs1 = New ADODB.Recordset

 ssql = ssql & " SELECT PC.CCTACOD"
 ssql = ssql & " FROM DBConsolidada..ProductoBloqueosConsol PC "
 ssql = ssql & " WHERE cMovNroDbl IS NULL AND nBlqMotivo = 3 And PC.cCtaCod Like '_____233%' "
 ssql = ssql & "        and Convert(Date,dCierre) = '" & Format(pdFecha, "YYYY/MM/DD") & "'"
 Set oRs1 = oCon.Execute(ssql)
 If (oRs1.BOF Or oRs1.EOF) Then
    If Mid(psCtaCod, 6, 3) = "232" Then
       If pbInactiva = 1 Then
          psCtaContCod = "21" & Mid(Trim(psCtaCod), 9, 1) & "202"
       Else
          psCtaContCod = "21" & Mid(Trim(psCtaCod), 9, 1) & "201"
       End If
    ElseIf Mid(psCtaCod, 6, 3) = "233" Then
          psCtaContCod = "21" & Mid(Trim(psCtaCod), 9, 1) & "303"
    ElseIf Mid(psCtaCod, 6, 3) = "234" Then
          psCtaContCod = "21" & Mid(Trim(psCtaCod), 9, 1) & "305"
    End If
    If (pnPersoneria = 1 Or pnPersoneria = 2) And Mid(psCtaCod, 6, 3) <> "234" Then
        psCtaContCod = psCtaContCod & "01"
    End If
    If (pnPersoneria = 2 Or pnPersoneria = 3) Then
        psCtaContCod = psCtaContCod & "02"
    Else
        If Mid(psCtaCod, 6, 3) <> "234" Then
            psCtaContCod = psCtaContCod & "01"
        Else
            If psCodInst = "1090100012521" Then
                psCtaContCod = psCtaContCod & "01"
            Else
                psCtaContCod = psCtaContCod & "02"
            End If
        End If
    End If
    psCtaContCod = psCtaContCod & Mid(Trim(psCtaCod), 4, 2)
 Else
    psCtaContCod = "21" & Mid(Trim(psCtaCod), 9, 1) & "70401"
    psCtaContCod = psCtaContCod & "01"
    psCtaContCod = psCtaContCod & Mid(Trim(psCtaCod), 4, 2)
 End If
 CreaCuentaAhorroDN = psCtaContCod
End Function
'**Creado por DAOR 20100623, Proyecto Basilea II ***********************************
Public Function VarAsientoEquivEmpSistFinanc(ByVal psCtaCod As String, oCon As ADODB.Connection) As String
Dim lrs As New ADODB.Recordset
Dim lsSQL As String
      
    sNS = "":   sIF = ""
    
    lsSQL = "exec B2_stp_sel_EquivalenciaEmpSistFinanciero '" & psCtaCod & "'"
    
    Set lrs = oCon.Execute(lsSQL)
    If Not (lrs.BOF Or lrs.EOF) Then
        sNS = IIf(IsNull(lrs!cEquivNS), "", Trim(lrs!cEquivNS))
        sIF = IIf(IsNull(lrs!cEquivIF), "", Trim(lrs!cEquivIF))
    End If
    
    lrs.Close
    Set lrs = Nothing

    VarAsientoEquivEmpSistFinanc = sNS & sIF
    
End Function



Public Function VarAsientoPrestamosAdm(ByVal sCuenta As String, oCon As ADODB.Connection) As String
    Dim RegCta As ADODB.Recordset
    Dim tmpSql As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    Set RegCta = New ADODB.Recordset
    'RegCta.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText 'ARCV 26-06-2007
    'oCon.AbreConexion
    tmpSql = "Select C.* from ColocGarantia C JOIN Garantias G ON C.cNumGarant = G.cNumGarant " _
        & "Where C.cCtaCod IN ('" & sCuenta & "') And C.nEstado = 1 And G.nEstado IN (3,6) " _
        & "And G.cTpoDoc = '121'"
    
    'RegCta.Open tmpSql, oCon, adOpenStatic, adLockReadOnly, adCmdText 'ARCV 26-06-2007
    '**ALPA***20080906
    Set RegCta = oCon.Execute(tmpSql)
    'Set RegCta = oCon.CargaRecordSet(tmpSql)
    If (RegCta.BOF Or RegCta.EOF) Then
        VarAsientoPrestamosAdm = "02"
    Else
        VarAsientoPrestamosAdm = "01"
    End If
    RegCta.Close
    Set RegCta = Nothing
End Function

Public Function CierreRealizado(Optional pnTipo As Integer = 1) As Boolean
Dim rsVarSis As ADODB.Recordset
Dim pdCieDia As Date
Dim pdCieMes As Date
Dim oCon As DConecta
Set oCon = New DConecta
Set rsVarSis = New ADODB.Recordset
Dim Sql As String

oCon.AbreConexion

rsVarSis.CursorLocation = adUseClient
Sql = "select nConsSisCod,nConsSisDesc, nConsSisValor From ConstSistema where nConsSisCod In (" & ConstSistemas.gConstSistCierreSistema & "," & ConstSistemas.gConstSistCierreMesNegocio & ")"
Set rsVarSis = oCon.CargaRecordSet(Sql)
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
    CierreRealizado = IIf(pdCieDia = gdFecSis, True, False)
ElseIf pnTipo = 2 Then
    CierreRealizado = IIf(pdCieMes = gdFecSis, True, False)
End If
End Function
Public Function CierreRealizado2(Optional pnTipo As Integer = 1, Optional conecc As ADODB.Connection) As Boolean
Dim rsVarSis As ADODB.Recordset
Dim pdCieDia As Date
Dim pdCieMes As Date
'Dim oCon As DConecta
'Set oCon = New DConecta
Set rsVarSis = New ADODB.Recordset
Dim Sql As String

'oCon.AbreConexion

'rsVarSis.CursorLocation = adUseClient
Sql = "select nConsSisCod,nConsSisDesc, nConsSisValor From ConstSistema where nConsSisCod In (" & ConstSistemas.gConstSistCierreSistema & "," & ConstSistemas.gConstSistCierreMesNegocio & ")"
'ALPA 20080611
Set rsVarSis = conecc.Execute(Sql)
'Set rsVarSis = oCon.CargaRecordSet(sql)
'Set rsVarSis.ActiveConnection = Nothing
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
    CierreRealizado2 = IIf(pdCieDia = gdFecSis, True, False)
ElseIf pnTipo = 2 Then
    CierreRealizado2 = IIf(pdCieMes = gdFecSis, True, False)
End If
End Function

Public Function VerificaDiaHabil(ByVal pdFecha As Date, pnTipo As Integer) As Boolean
     Dim lsFchPro As String
        Dim ldFchPro As Date
        Dim lnMes As Integer
        Dim ldAno As Date
        '------------ DCTO DE INACTIVAS CMAC-CUSCO-------
        Dim oCapG As COMDCaptaGenerales.DCOMCaptaGenerales
        Dim rs As New ADODB.Recordset
        '------------------------------------------------
        ldAno = pdFecha
        lnMes = Month(pdFecha)
        If lnMes = 12 Then
            lnMes = 0
            ldAno = DateAdd("yyyy", 1, ldAno)
        End If
    Select Case pnTipo
        Case 1
        '   Busca Primer día del mes
            lsFchPro = "01" & "/" & FillNum(Trim(Str(lnMes + 1)), 2, "0") & "/" & Trim(Str(Year(ldAno)))
            ldFchPro = CDate(lsFchPro)
        Case 2
        '   Busca día anterior al último día del mes para procesar descuento por Inactivas
        '------------------ DCTO INACTIVAS CMAC CUSCO -----------------
            Set oCapG = New COMDCaptaGenerales.DCOMCaptaGenerales
                Set rs = oCapG.ObtenerDctoInactivasFecha()
            Set oCapG = Nothing
            If rs Is Nothing Then
            Else
                Do Until rs.EOF
                   lsFchPro = rs!nDia & "/" & rs!nMes & "/" & Trim(Str(Year(ldAno)))
                   ldFchPro = CDate(lsFchPro)
                   If pdFecha = ldFchPro Then
                        Exit Do
                   End If
                   rs.MoveNext
                Loop
            End If
        '------------------------------------------------------------------
        
'            lsFchPro = "01" & "/" & FillNum(Trim(Str(lnMes + 1)), 2, "0") & "/" & Trim(Str(Year(ldAno)))
'            ldFchPro = DateAdd("d", -2, CDate(lsFchPro))
        Case 3
        '   Busca ultimo día del mes para procesar Cierre de Mes
            lsFchPro = "01" & "/" & FillNum(Trim(Str(lnMes + 1)), 2, "0") & "/" & Trim(Str(Year(ldAno)))
            ldFchPro = DateAdd("d", -1, CDate(lsFchPro))
    End Select
    If pdFecha = ldFchPro Then
        VerificaDiaHabil = True
    Else
        VerificaDiaHabil = False
    End If

End Function

Public Function ValdiaText(pnKey As Integer) As Integer
    If pnKey = 8 Or (Chr(pnKey) >= "A" And Chr(pnKey) <= "Z") Or (Chr(pnKey) >= "a" And Chr(pnKey) <= "z") Or (Chr(pnKey) >= "0" And Chr(pnKey) <= "9") Or Chr(pnKey) = " " Or Chr(pnKey) = "/" Or Chr(pnKey) = "\" Or Chr(pnKey) = ":" Then
        ValdiaText = pnKey
    Else
        ValdiaText = 0
    End If
End Function
'--- LAYG 07/01/2005 - ica
Public Function VarAG(ByVal pCuentaProd As String, ByVal pCtaCont As String, Optional ByVal oCon As ADODB.Connection = Nothing) As String
Dim lsAge As String
    '******** ica parche en caja agencias EJRS 13/10/2004 *****************************
    If Left(pCtaCont, 2) = "11" Then
        lsAge = Mid(pCuentaProd, 4, 2)
    Else
        lsAge = Mid(pCuentaProd, 4, 2)
        
        '*** PEAC 20140204
        Select Case lsAge
            Case "08"
                        
                Dim rs As ADODB.Recordset
                Dim Sql As String
                

                Set rs = New ADODB.Recordset
                Sql = "exec stp_sel_BuscaAgeColocaciones '" & pCuentaProd & "' "

                Set rs = oCon.Execute(Sql)
                
                If Not rs.EOF And Not rs.BOF Then
                    lsAge = Trim(rs!cAgeCodAct)
                End If
                rs.Close
                Set rs = Nothing

        End Select
        '*** FIN PEAC
        
        'Select Case Mid(pCuentaProd, 4, 2)
        '    Case "02" 'Chincha
        '        'lsAge = "01"
        '        'lsAge = "02"
        '    Case "03" 'Imperail
        '        'lsAge = "03"
        '        'lsAge = "13"
        '    Case "04"
        '        'lsAge = "04"
        '        'lsAge = "04"
         '   Case "05"
        '        'lsAge = "04"
        '    Case "06"
        '        'lsAge = "03"
        '    Case "07"
        '        'lsAge = "05"
        '    Case "08"
        '        'lsAge = "06"
        '    Case "13"
        '        'lsAge = "03"
        '    Case Else
        '        lsAge = Mid(pCuentaProd, 4, 2)
        'End Select
    End If
    '**** CULMINACION DE PARCHES JEAN 13/10/2004 ***********+
    VarAG = lsAge
End Function

Public Function GetPlantillaPuente(ByVal psPlantillaAnt As String, ByVal lsProducto As String, ByVal pnConcepto As Long, ByVal psOpecod As String, oCon As ADODB.Connection) As String
''''Dim oCon As DConecta
'''Dim rs As ADODB.Recordset
'''Dim sql As String
''''Set oCon = New DConecta
'''
'''GetPlantillaPuente = ""
'''Set rs = New ADODB.Recordset
''''oCon.AbreConexion
'''sql = "SELECT   cCtaContCodNew FROM PlantillaAsientoPuente " _
'''    & " where   cCtaContCodAnt = '" & psPlantillaAnt & "' and cProducto ='" & lsProducto & "' " _
'''    & "         and nPrdConceptoCod = " & pnConcepto & " AND cOpeCod ='" & psOpeCod & "'"
'''rs.Open sql, oCon, adOpenStatic, adLockReadOnly, adCmdText
'''
'''If Not rs.EOF And Not rs.BOF Then
'''    GetPlantillaPuente = Trim(rs!cCtaContCodNew)
'''End If
'''rs.Close
'''Set rs = Nothing
'ALPA**20080604
Dim rs As ADODB.Recordset
Dim Sql As String
'Set oCon = New DConecta

GetPlantillaPuente = ""
Set rs = New ADODB.Recordset
'oCon.AbreConexion
Sql = "SELECT   cCtaContCodNew FROM PlantillaAsientoPuente " _
    & " where   cCtaContCodAnt = '" & psPlantillaAnt & "' and cProducto ='" & lsProducto & "' " _
    & "         and nPrdConceptoCod = " & pnConcepto & " AND cOpeCod ='" & psOpecod & "'"
'ALPA***ASIENTO
Set rs = oCon.Execute(Sql)
'rs.Open sql, oCon, adOpenStatic, adLockReadOnly, adCmdText

If Not rs.EOF And Not rs.BOF Then
    GetPlantillaPuente = Trim(rs!cCtaContCodNew)
End If
rs.Close
Set rs = Nothing
End Function
Public Function ValidaOk2(ByVal pFecha As Date, Optional pNuePlan As Boolean = False, Optional psAgeCod As String = "", Optional conecc As ADODB.Connection) As String
Dim vDifere As Currency, vMovAsi As Currency, vMovSal As Currency
Dim vSHoy As Currency, vSAyer As Currency
Dim vDHoy As Currency, vDAyer As Currency
Dim vFecha As Date
Dim nDias As Integer
Dim x As Integer
Dim lsCtaCaja As String
'Dim oCon As DConectaA
nDias = 30 'Val(ReadVarSis("ADM", "nDiaValAsiento"))
x = 0
ValidaOk2 = "" ':  vCieDia = 0
'********************************************************************************************
'CTA. 111103AG   --->  A  Cajaff
vFecha = pFecha
'vSHoy = Billetaje(vFecha, "1", pFecha)
'vDHoy = Billetaje(vFecha, "2", pFecha)

vSHoy = Billetaje2(vFecha, "1", psAgeCod, conecc)
vDHoy = Billetaje2(vFecha, "2", psAgeCod, conecc)
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
    vSAyer = Billetaje2(vFecha, "1", psAgeCod, conecc)
    vDAyer = Billetaje2(vFecha, "2", psAgeCod, conecc)

End If
'Verifica que cuadre la cta. 111103AG o 111102AG (NuevoPlan) (Nela 17.09.2001)
'Set oCon = New DConectaA
'oCon.AbreConexion

vMovSal = vSHoy - vSAyer
vMovAsi = CtaAsiento(pFecha, "A", "1", "1", pNuePlan, psAgeCod, conecc)
vDifere = (vMovSal - vMovAsi)
lsCtaCaja = AsientoParche("111102" & Right(gsCodAge, 2), True, conecc)
If lsCtaCaja = "" Then
    lsCtaCaja = "111102" & Right(gsCodAge, 2)
End If



If (vDifere < 0 And vDifere >= -0.05) Then
        'ARCV 31-03-2007
        'sSql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','63110909'," & Abs(vDifere) & ",0,'0','" & psAgeCod & "') "
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','421229" & Right(gsCodAge, 2) & "'," & Abs(vDifere) & ",0,'0','" & psAgeCod & "') "
        '-------------
        'oCon.Ejecutar sSQL
        conecc.Execute (ssql)
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','" & lsCtaCaja & "' ,0," & Abs(vDifere) & ",'0','" & psAgeCod & "') "
        'oCon.Ejecutar sSQL
        conecc.Execute (ssql)
ElseIf (vDifere > 0 And vDifere <= 0.05) Then
      'ARCV 31-03-2007
      'ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','5212290299" & Right(gsCodAge, 2) & "' ,0," & Abs(vDifere) & ",'0','" & psAgeCod & "')"
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','521229" & Right(gsCodAge, 2) & "' ,0," & Abs(vDifere) & ",'0','" & psAgeCod & "')"
        'oCon.Ejecutar sSQL
        conecc.Execute (ssql)
      '---------------
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo, cCodAge) " & _
            " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','" & lsCtaCaja & "' ," & Abs(vDifere) & ",0,'0','" & psAgeCod & "') "
        'oCon.Ejecutar sSQL
        conecc.Execute (ssql)
End If
'**
If Abs(vDifere) > 0.05 Then
' VALIDACIONES PARA CTA 11 (SILVITA - NELA) ?????
    If pNuePlan Then
        ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "     * No cuadra la cta.cnt. 111102AG en Soles, diferencia de " & Str(vDifere)
    Else
        ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "     * No cuadra la cta.cnt. 111103AG en Soles, diferencia de " & Str(vDifere)
    End If
    ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If

vMovSal = vDHoy - vDAyer
'vMovAsi = CtaAsiento(pFecha, "A", "2", "2", pNuePlan, psAgeCod, oCon.ConexionActiva)
vMovAsi = CtaAsiento(pFecha, "A", "2", "2", pNuePlan, psAgeCod, conecc)
vDifere = (vMovSal - vMovAsi)
'lsCtaCaja = AsientoParche("112102" & Right(gsCodAge, 2), True, oCon.ConexionActiva)
lsCtaCaja = AsientoParche("112102" & Right(gsCodAge, 2), True, conecc)
If lsCtaCaja = "" Then
    lsCtaCaja = "112102" & Right(gsCodAge, 2)
End If

'**
If (vDifere < 0 And vDifere >= -0.05) Then
      'ARCV 31-03-2007
      'sSql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','63210909'," & Abs(vDifere) & ",0,'0')"
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','422229" & Right(gsCodAge, 2) & "'," & Abs(vDifere) & ",0,'0')"
        '--------
        'oCon.Ejecutar sSQL
        conecc.Execute (ssql)
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','" & lsCtaCaja & "' ,0," & Abs(vDifere) & ",'0')"
        'oCon.Ejecutar sSQL
        conecc.Execute (ssql)
ElseIf (vDifere > 0 And vDifere <= 0.05) Then
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','522229" & Right(gsCodAge, 2) & " ',0," & Abs(vDifere) & ",'0')"
        'oCon.Ejecutar sSQL
        conecc.Execute (ssql)
        ssql = "INSERT INTO AsientoDN (dfecha, cctacnt, ndebe, nhaber, ctipo) " & _
           " VALUES('" & Format(vFecha, "mm/dd/yyyy hh:mm:ss") & "','" & lsCtaCaja & "'," & Abs(vDifere) & ",0,'0')"
        'oCon.Ejecutar sSQL
        conecc.Execute (ssql)
End If
'**
If Abs(vDifere) > 0.05 Then
    If pNuePlan Then
        ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "     * No cuadra la cta.cnt. 112102AG en Dólares, diferencia de " & Str(vDifere)
    Else
        ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "     * No cuadra la cta.cnt. 112103AG en Dólares, diferencia de " & Str(vDifere)
    End If
    ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
'********************************************************************************************
'CTA. 14 - (1419)   ---> B   Créditos
'Verifica que cuadre la cta. 14
vMovSal = Abs(EstadCred2(pFecha, "1", 1, psAgeCod, conecc) + EstadCred2(pFecha, "1", 2, psAgeCod, conecc))
'vMovAsi = Abs(CtaAsiento(pFecha, "B", "1", , pNuePlan, psAgeCod, oCon.ConexionActiva))
vMovAsi = Abs(CtaAsiento(pFecha, "B", "1", , pNuePlan, psAgeCod, conecc))
vDifere = (vMovSal - vMovAsi)
If vDifere <> 0 Then
    ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "     * No cuadra la cta.cnt. 14 en Soles, diferencia de " & Str(vDifere)
    ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
vMovSal = Abs(EstadCred2(pFecha, "2", 1, psAgeCod, conecc))
vMovAsi = Abs(CtaAsiento(pFecha, "B", "2", , pNuePlan, psAgeCod, conecc))
vDifere = (vMovSal - vMovAsi)
If vDifere <> 0 Then
    ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "     * No cuadra la cta.cnt. 14 en Dólares, diferencia de " & Str(vDifere)
    ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
'********************************************************************************************
'CTA. 23,24 y 26  ---> C  Ahorros
'Verifica que cuadren las ctas. 23,24 y 26 o 21 (NuevoPlan)
vMovSal = Abs(EstadAho2(pFecha, "1", 1, psAgeCod, conecc) + EstadAho2(pFecha, "1", 2, psAgeCod, conecc) + EstadAho2(pFecha, "1", 3, psAgeCod, conecc))
vMovAsi = Abs(CtaAsiento(pFecha, "C", "1", , pNuePlan, psAgeCod, conecc))
vDifere = Round((vMovSal - vMovAsi), 2)
If vDifere <> 0 Then
    If pNuePlan Then
        ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "     * No cuadran las ctas.cnts. 2112, 2113, 2312 y 2313 en Soles, diferencia de " & Str(vDifere)
    Else
        ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "     * No cuadran las ctas.cnts. 23, 24 y 26 en Soles, diferencia de " & Str(vDifere)
    End If
    ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
vMovSal = Abs(EstadAho2(pFecha, "2", 1, psAgeCod, conecc) + EstadAho2(pFecha, "2", 2, psAgeCod, conecc) + EstadAho2(pFecha, "2", 3, psAgeCod, conecc))
vMovAsi = Abs(CtaAsiento(pFecha, "C", "2", , pNuePlan, psAgeCod, conecc))
vDifere = (vMovSal - vMovAsi)
If vDifere <> 0 Then
    If pNuePlan Then
        ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "     * No cuadran las ctas.cnts. 2122, 2123, 2322 y 2323 en Dólares, diferencia de " & Str(vDifere)
    Else
        ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "     * No cuadran las ctas.cnts. 23, 24 y 26 en Dólares, diferencia de " & Str(vDifere)
    End If
    ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "        - Movimiento de los Saldos " & Format(vMovSal, "#0.00")
    ValidaOk2 = ValidaOk2 & oImpresora.gPrnSaltoLinea & "        - Movimiento en el Asiento " & Format(vMovAsi, "#0.00")
End If
If Not ValidaOk2 = "" And Not psAgeCod = "" Then
    ValidaOk2 = Chr(10) & Chr(10) & "AGENCIA " & psAgeCod & ": " & ValidaOk2
End If
'oCon.CierraConexion
'Set oCon = Nothing
End Function
Public Function Billetaje2(ByVal pFecha As Date, ByVal pMoneda As String, Optional psAgeCod As String = "", Optional rsConeccion As ADODB.Connection) As Currency
    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
'    Dim oCon As DConectaA
'    Set oCon = New DConectaA
    
'    oCon.AbreConexion
    
'    tmpSql = " Select Sum(nMonto) Campo From Mov M" _
'           & " Inner Join MovUserEfectivo ME On M.nMovNro = ME.nMovNro" _
'           & " Where M.cMovNro Like '" & Format(pFecha, gsFormatoMovFecha) & "%' And ME.nMonto <> 0 " _
'           & " And ME.cEfectivoCod LIKE '" & pMoneda & "%' And M.cOpeCod IN ('901007','901016') " _
'           & "AND Substring(cMovNro, 18,2) in (SELECT distinct right(cCtaCnt,2) FROM AsientoDN " _
'           & "WHERE datediff(dd,dfecha,'" & Format(pFechaMov, "mm/dd/yyyy") & "') = 0 " _
'           & "AND Substring(cCtaCnt,1,6) IN ('11" & pMoneda & "102','11" & pMoneda & "103'))"

'ARCV 28-03-2007
'tmpSql = "Select ISNULL(SUM(E.nMonto),0) Campo From Mov M INNER JOIN MovUserEfectivo E ON M.nMovNro = E.nMovNro " _
    & "Where M.nMovFlag NOT IN (2,1)  And LEFT(M.cMovNro,8) IN (Select MAX(LEFT(M1.cMovNro,8)) From Mov M1 " _
    & "JOIN MovUserEfectivo E1 ON M1.nMovNro = E1.nMovNro Where M1.nMovFlag NOT IN (2,1) And E1.cEfectivoCod LIKE '" & pMoneda & "%' And " _
    & "LEFT(M1.cMovNro,8) <= '" & Format(pFecha, gsFormatoMovFecha) & "' AND M.cOpeCod in ('901007', '901016')) AND M.cOpeCod in ('901007', '901016') And E.cEfectivoCod LIKE '" & pMoneda & "%'"
 
tmpSql = "SELECT ISNULL(SUM(E.nMonto),0) Campo From Mov M INNER JOIN MovUserEfectivo E ON M.nMovNro = E.nMovNro " _
    & "Where M.nMovFlag NOT IN (2,1)  And LEFT(M.cMovNro,8) IN (Select MAX(LEFT(M1.cMovNro,8)) From Mov M1 " _
    & "JOIN MovUserEfectivo E1 ON M1.nMovNro = E1.nMovNro Where M1.nMovFlag NOT IN (2,1) And E1.cEfectivoCod LIKE '" & pMoneda & "%' And " _
    & "LEFT(M1.cMovNro,8) <= '" & Format(pFecha, gsFormatoMovFecha) & "' AND M.cOpeCod in ('901007', '901016')" _
    & IIf(psAgeCod = "", "", " AND Substring(M1.cMovNro,18,2) = '" & psAgeCod & "'") & ") AND M.cOpeCod in ('901007', '901016') And E.cEfectivoCod LIKE '" & pMoneda & "%'"
'----------

If Not psAgeCod = "" Then
    tmpSql = tmpSql & " and substring(cmovnro,18,2) = '" & psAgeCod & "'"
End If

    'Set tmpReg = oCon.CargaRecordSet(tmpSql)
    Set tmpReg = rsConeccion.Execute(tmpSql)
    If (tmpReg.BOF Or tmpReg.EOF) Then
        Billetaje2 = 0
    Else
        Billetaje2 = IIf(IsNull(tmpReg!Campo), 0, tmpReg!Campo)
    End If
    tmpReg.Close
    Set tmpReg = Nothing
'    Set oCon = Nothing
End Function
Public Function EstadCred2(ByVal pFecha As Date, ByVal pMoneda As String, ByVal pOption As Integer, Optional psAgeCod As String = "", Optional conecc As ADODB.Connection) As Currency
Dim tmpReg As ADODB.Recordset
Dim tmpSql As String
'Dim oCon As DConectaA
'Set oCon = New DConectaA

'oCon.AbreConexion

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
'tmpReg.CursorLocation = adUseClient
'Set tmpReg = oCon.CargaRecordSet(tmpSql)
Set tmpReg = conecc.Execute(tmpSql)
'Set tmpReg.ActiveConnection = Nothing
If (tmpReg.BOF Or tmpReg.EOF) Then
    EstadCred2 = 0
Else
    EstadCred2 = IIf(IsNull(tmpReg!Campo), 0, tmpReg!Campo)
End If
tmpReg.Close
Set tmpReg = Nothing
'Set oCon = Nothing
End Function
Public Function EstadAho2(ByVal pFecha As Date, ByVal pMoneda As String, ByVal pOption As Integer, Optional psAgeCod As String = "", Optional conecc As ADODB.Connection) As Currency
Dim tmpReg As ADODB.Recordset
Dim tmpSql As String
''Dim oCon As DConectaA
''Set oCon = New DConectaA

''oCon.AbreConexion

If pOption = 1 Then
    tmpSql = " SELECT (ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(DateAdd("d", -1, pFecha), gsFormatoFecha) & "') = 0 then nSaldo End)),0)" _
           & "         -  ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(pFecha, gsFormatoFecha) & "') = 0 then nSaldo End)),0)) AS Campo" _
           & " FROM CapEstadSaldo WHERE nMoneda = '" & pMoneda & "' And nProducto = " & Producto.gCapAhorros & _
        IIf(psAgeCod = "", "", " and cCodAge = '" & psAgeCod & "'")
ElseIf pOption = 2 Then
    tmpSql = " SELECT    (ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(DateAdd("d", -1, pFecha), gsFormatoFecha) & "') = 0 then nSaldo End)),0)" _
           & "         -  ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(pFecha, gsFormatoFecha) & "') = 0 then nSaldo End)),0)) AS Campo" _
           & " FROM CapEstadSaldo WHERE nMoneda = '" & pMoneda & "' And nProducto = " & Producto.gCapPlazoFijo & _
        IIf(psAgeCod = "", "", " and cCodAge = '" & psAgeCod & "'")
Else
    tmpSql = " SELECT    (ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(DateAdd("d", -1, pFecha), gsFormatoFecha) & "') = 0 then nSaldo End)),0)" _
           & "         -  ISNULL((SUM(CASE when datediff(dd,dEstad,'" & Format(pFecha, gsFormatoFecha) & "') = 0 then nSaldo End)),0)) AS Campo" _
           & " FROM CapEstadSaldo WHERE nMoneda = '" & pMoneda & "' And nProducto = " & Producto.gCapCTS & _
        IIf(psAgeCod = "", "", " and cCodAge = '" & psAgeCod & "'")
End If
Set tmpReg = New ADODB.Recordset
'tmpReg.CursorLocation = adUseClient
'Set tmpReg = oCon.CargaRecordSet(tmpSql)
Set tmpReg = conecc.Execute(tmpSql)
'Set tmpReg.ActiveConnection = Nothing
If (tmpReg.BOF Or tmpReg.EOF) Then
    EstadAho2 = 0
Else
    EstadAho2 = IIf(IsNull(tmpReg!Campo), 0, tmpReg!Campo)
End If
tmpReg.Close
Set tmpReg = Nothing
End Function
'PASI20161202 CCE********************************************************************************************************
Public Function EsMovDeChequeCCE(ByVal pnMovNro As Long, Optional oCon As ADODB.Connection = Nothing) As Boolean
Dim ssql As String
Dim rs As ADODB.Recordset
    ssql = "stp_sel_CCE_EsMovdeCHequeCCE " & pnMovNro
  Set rs = oCon.Execute(ssql)
  If Not (rs.EOF And rs.BOF) Then
    EsMovDeChequeCCE = IIf(rs!nCant = 0, False, True)
  Else
    EsMovDeChequeCCE = False
  End If
End Function
'PASI END****************************************************************************************************************
'VAPA20170202 CCE********************************************************************************************************
Public Function EsOpeCodCheCCE(Optional psOpeCodChe As String = "", Optional oCon As ADODB.Connection = Nothing) As Boolean
Dim ssql As String
Dim rs As ADODB.Recordset
    ssql = "stp_sel_CCE_EsOpeCodChe " & psOpeCodChe
  Set rs = oCon.Execute(ssql)
  If Not (rs.EOF And rs.BOF) Then
    EsOpeCodCheCCE = IIf(rs!cVal = "no", False, True)
  Else
    EsOpeCodCheCCE = False
  End If
End Function
'VAPA END
'VAPA20170206 CCE********************************************************************************************************
Public Function EsOpeCodCambioACCE(ByVal pnMovNro As Long, Optional psOpeCodChe As String = "", Optional ByVal pnMontoDebe As Double = 0, Optional oCon As ADODB.Connection = Nothing) As Boolean
Dim ssql As String
Dim rs As ADODB.Recordset
    ssql = "stp_sel_CCE_EsConcepto3 '" & pnMovNro & "','" & psOpeCodChe & "','" & pnMontoDebe & "'"
  Set rs = oCon.Execute(ssql)
  If Not (rs.EOF And rs.BOF) Then
    EsOpeCodCambioACCE = IIf(rs!bVal = 0, False, True)
  Else
    EsOpeCodCambioACCE = False
  End If
End Function
'VAPA END

