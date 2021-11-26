VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DbalanceCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sConexion As String
Dim sCentralCom As String

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As New ClasIni
   sCentralCom = oIni.BaseComunes
Set oIni = Nothing
Set dbConec = New DConecta
If Not dbConec.AbreConexion() Then
   Call RaiseError(MyUnhandledError, "NBalanceCont:Initialize Method. " & Err.Description)
End If
End Sub

Private Sub Class_Terminate()
dbConec.CierraConexion
Set dbConec = Nothing
End Sub

Public Sub EliminaBalance(pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer, Optional lbEjecutarBatch As Boolean = False)
On Error GoTo EliminaBalanceErr
   psSql = "DELETE BalanceEstad " _
        & "WHERE cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
        & "' and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
        & "cBalanceAnio = '" & Format(pnAnio, "0000") & "'"
   If lbEjecutarBatch Then
      dbConec.AdicionaCmdBatch psSql
   Else
      dbConec.Ejecutar psSql
   End If
Exit Sub
EliminaBalanceErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:EliminaBalance Method")
End Sub
Public Sub EliminaBalanceFormaAB(psTipoRepoFormula As Integer, psFecha As String, Optional lbEjecutarBatch As Boolean = False)
On Error GoTo EliminaABBalanceErr
   psSql = "DELETE BalanceGen WHERE cFecha = '" & psFecha & "' And cBalanceCate = '" & psTipoRepoFormula & "'"
   If lbEjecutarBatch Then
      dbConec.AdicionaCmdBatch psSql
   Else
      dbConec.Ejecutar psSql
   End If
Exit Sub
EliminaABBalanceErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:EliminaABBalance Method")
End Sub
    

Public Sub EliminaBalanceTemp(pnTipoBala As Integer, pnMoneda As Integer, Optional lbEjecutarBatch As Boolean = False)
On Error GoTo EliminaBalanceErr
   psSql = "DELETE BalanceTemp " _
        & "WHERE cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda & "'"
   If lbEjecutarBatch Then
      dbConec.AdicionaCmdBatch psSql
   Else
      dbConec.Ejecutar psSql
   End If
Exit Sub
EliminaBalanceErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:EliminaBalance Method")
End Sub

Public Function ValidaSaldosIniciales(pnMes As Integer, pnAnio As Integer) As String
Dim lnMes As Integer
Dim lnAnio As Integer
Dim lsFecha As String
Dim lsMsg As String

lnMes = pnMes - 1
lnAnio = pnAnio
If pnMes = 1 Then
   lnMes = 12
   lnAnio = pnAnio - 1
End If
lsFecha = "01/" & Format(pnMes, "00") & "/" & Format(pnAnio, "0000")
lsFecha = Format(CDate(lsFecha) - 1, gsFormatoFecha)
ValidaSaldosIniciales = ""

If pnMes = 1 Then
    psSql = "select ISNULL(b1.cCtaContCod,b2.cCtaContCod), nSaldoIniImporte, nSaldoFinImporte " _
          & " From ( Select cs.cCtaContCod, nCtaSaldoImporte nSaldoIniImporte " _
          & "        from CtaSaldo cs, CtaCont cc " _
          & " where dCtaSaldoFecha = (Select Max(dCtaSaldoFecha) FROM CtaSaldo cs1 WHERE cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha < = '" & lsFecha & "') " _
          & "              and cc.cCtaContCod like cs.cCtaContCod + '%' " _
          & "        group by cs.cCtaContCod, nCtaSaldoImporte " _
          & "        having count(*) = 1 " _
          & "      ) b2 LEFT JOIN " _
          & "      ( " _
          & "        Select cCtaContCod, nSaldoFinImporte from BalanceEstad " _
          & "        where cBalanceCate = 2 and cBalanceTipo in (1,2) and cBalanceMes = '" & Format(lnMes, "00") & "' and cBalanceAnio = '" & Format(lnAnio, "0000") & "' And cCtaContCod Not Like '[4563]%' " _
          & "        Union   " _
          & "        Select cCtaContCod, nSaldoFinImporte from BalanceEstad " _
          & "        where cBalanceCate = 2 and cBalanceTipo in (1,2) and cBalanceMes = '" & Format(lnMes + 1, "00") & "' and cBalanceAnio = '" & Format(lnAnio, "0000") & "' And cCtaContCod Like '[4563]%'" _
          & "      ) b1 ON b1.cCtaContCod = b2.cCtaContCod " _
          & "      Where IsNull(b1.nSaldoFinImporte, 0) <> IsNull(b2.nSaldoIniImporte, 0) "
Else
    psSql = "select ISNULL(b1.cCtaContCod,b2.cCtaContCod), nSaldoIniImporte, nSaldoFinImporte " _
          & " From ( Select cs.cCtaContCod, nCtaSaldoImporte nSaldoIniImporte " _
          & "        from CtaSaldo cs, CtaCont cc " _
          & " where dCtaSaldoFecha = (Select Max(dCtaSaldoFecha) FROM CtaSaldo cs1 WHERE cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha < = '" & lsFecha & "') " _
          & "              and cc.cCtaContCod like cs.cCtaContCod + '%' " _
          & "        group by cs.cCtaContCod, nCtaSaldoImporte " _
          & "        having count(*) = 1 " _
          & "      ) b2 LEFT JOIN " _
          & "      ( " _
          & "        Select cCtaContCod, nSaldoFinImporte from BalanceEstad " _
          & "        where cBalanceCate = 2 and cBalanceTipo in (1,2) and cBalanceMes = '" & Format(lnMes, "00") & "' and cBalanceAnio = '" & Format(lnAnio, "0000") & "' " _
          & "      ) b1 ON b1.cCtaContCod = b2.cCtaContCod " _
          & "      Where IsNull(b1.nSaldoFinImporte, 0) <> IsNull(b2.nSaldoIniImporte, 0) "
End If

Set prs = dbConec.CargaRecordSet(psSql)
If Not prs.EOF Then
   ValidaSaldosIniciales = "Se detectaron las siguiente Cuentas con diferencias en sus Saldos Iniciales" & Chr(13) _
         & prs.GetString()
End If
End Function


Public Sub InsertaSaldosIniciales(pnTipoBala As Integer, pnMoneda As Integer, psFecha As String, Optional lbEjecutarBatch As Boolean = False)
Dim sCondBala As String
Dim sCta      As String
Dim sCtaMN    As String
On Error GoTo InsertaSaldosInicialesErr
Set prs = dbConec.CargaRecordSet("SELECT cCtaContCod FROM CtaContClase WHERE nCtaMoneda = 1")
Do While Not prs.EOF
   sCtaMN = sCtaMN & "'" & prs!cCtaContCod & "',"
   prs.MoveNext
Loop
RSClose prs
If sCtaMN <> "" Then
   sCtaMN = Left(sCtaMN, Len(sCtaMN) - 1)
End If
If pnMoneda <> 0 And pnMoneda <> 9 Then
   If pnMoneda = 1 Then
      sCondBala = "and ( (Substring(a.cCtaContCod,3,1) = '1' or substring(a.cCtaContCod,1,1) IN (" & sCtaMN & " ) " _
                & " or substring(a.cCtaContCod,1,2) IN (" & sCtaMN & " ) ) " _
                & IIf(pnTipoBala = 2, "or  substring(a.cCtaContCod,3,1) = '6' )", " and substring(a.cCtaContCod,3,1) <> '6' )")
   ElseIf pnMoneda = 2 Then
      sCondBala = "and ( (Substring(a.cCtaContCod,3,1) = '2' and NOT substring(a.cCtaContCod,1,1) IN (" & sCtaMN & ") " _
                & " and NOT substring(a.cCtaContCod,1,2) IN (" & sCtaMN & ") )" _
                & " and substring(a.cCtaContCod,3,1) <> '6' )"
   Else
      sCondBala = "and ( Substring(a.cCtaContCod,3,1) = '" & pnMoneda & "' " _
                & IIf(pnTipoBala = 2, "or  substring(a.cCtaContCod,3,1) = '6' )", " ) ")
   End If
   sCta = "a.cCtaContCod"
ElseIf pnMoneda = 9 Then
   sCta = "a.cCtaContCod"
   sCondBala = " and (Substring(a.cCtaContCod,3,1) IN ('1','2','3','4','5') " & IIf(pnTipoBala = 2, " or substring(a.cCtaContCod,3,1) = '6' )", ") ")
Else
   sCta = "SubString(a.cCtaContCod,1,2) + '0' + SubString(a.cCtaContCod,4,20) "
   sCondBala = " and (Substring(a.cCtaContCod,3,1) IN ('1','2','3','4','5') " & IIf(pnTipoBala = 2, " or substring(a.cCtaContCod,3,1) = '6' )", ") ")
End If

   psSql = "INSERT BALANCETEMP " _
        & "SELECT '" & pnTipoBala & "','" & pnMoneda & "'," & sCta & ", Sum(nCtaSaldoImporte), 0, 0, cCtaCaracter, '1' " _
        & "FROM CtaSaldo a JOIN CtaCont c ON c.cCtaContCod = a.cCtaContCod JOIN " & sCentralCom & "CtaContClase cls ON a.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' " _
        & "WHERE (nCtaSaldoImporte <> 0 or len(a.cCtaContCod) <=2) " & sCondBala & " and " _
        & "                                dCtaSaldoFecha = (SELECT MAX(dCtaSaldoFecha) FROM CtaSaldo b " _
        & "                                                  WHERE b.cCtaContCod = a.cCtaContCod and b.dCtaSaldoFecha < '" & psFecha & "')  " _
        & "GROUP BY " & sCta & ", cls.cCtaCaracter"
   dbConec.CommandTimeout 3000
   dbConec.Ejecutar psSql
Exit Sub
InsertaSaldosInicialesErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaSaldosIniciales Method")
End Sub

Public Sub InsertaMovimientosMes(pnTipoBala As Integer, pnMoneda As Integer, psFechaDel As String, psFechaAl As String, Optional pbBCient As Boolean = False, Optional pbConCierreAnual As Boolean = False)
Dim sCondBala As String
Dim sCta      As String
Dim sCtaMN    As String
On Error GoTo InsertaMovimientosMesErr
sCtaMN = "SELECT cCtaContCod FROM " & sCentralCom & "CtaContClase WHERE nCtaMoneda = 1"
If pbBCient Then
   If pnMoneda = 0 Then
      sCta = "substring(mc.cCtaContCod,1,2) + '0' + substring(mc.cCtaContCod,4,20)"
      sCondBala = ""
   ElseIf pnMoneda = 3 And pnTipoBala <> 5 Then
      sCta = "mc.cCtaContCod"
      sCondBala = " "
   Else
      sCondBala = " and (SubString(mc.cCtaContCod,3,1) = '" & pnMoneda & "') "
      sCta = "mc.cCtaContCod"
   End If
Else
   If pnMoneda <> 0 And pnMoneda <> 9 Then
      If pnMoneda = 1 Then
         sCondBala = "and ( (Substring(mc.cCtaContCod,3,1) = '1' or substring(mc.cCtaContCod,1,1) IN (" & sCtaMN & ") " _
                   & " or substring(mc.cCtaContCod,1,2) IN (" & sCtaMN & ") ) " _
                   & IIf(pnTipoBala = 2, "or  substring(mc.cCtaContCod,3,1) = '6' )", " and substring(mc.cCtaContCod,3,1) <> '6' )")
      ElseIf pnMoneda = 2 Then
         sCondBala = "and ( (Substring(mc.cCtaContCod,3,1) = '2' and NOT substring(mc.cCtaContCod,1,1) IN (" & sCtaMN & ") " _
                   & " and NOT substring(mc.cCtaContCod,1,2) IN (" & sCtaMN & ") )" _
                   & " and substring(mc.cCtaContCod,3,1) <> '6' )"
      Else
         sCondBala = "and ( Substring(mc.cCtaContCod,3,1) = '" & pnMoneda & "' " _
                   & IIf(pnTipoBala = 2, "or  substring(mc.cCtaContCod,3,1) = '6' )", " )")
      End If
      sCta = "mc.cCtaContCod"
   ElseIf pnMoneda = 9 Then
      sCta = "mc.cCtaContCod"
      sCondBala = " and (Substring(mc.cCtaContCod,3,1) IN ('1','2','3','4','5') " & IIf(pnTipoBala = 2, " or substring(mc.cCtaContCod,3,1) = '6' )", ") ")
   Else
      sCta = "SubString(mc.cCtaContCod,1,2) + '0' + SubString(mc.cCtaContCod,4,20) "
      sCondBala = " and (Substring(mc.cCtaContCod,3,1) IN ('1','2','3','4','5') " & IIf(pnTipoBala = 2, " or substring(mc.cCtaContCod,3,1) = '6' )", ") ")
   End If
End If
   psSql = "INSERT BALANCETEMP " _
        & "SELECT '" & pnTipoBala & "','" & pnMoneda & "'," & sCta & ", 0," _
        & "       SUM( ISNULL( CASE WHEN mc.nMovImporte > 0 THEN mc.nMovImporte " _
        & "                         WHEN mc.nMovImporte < 0 THEN 0 END, 0) ) as nDebe, " _
        & "       SUM( ISNULL( CASE WHEN mc.nMovImporte < 0 THEN mc.nMovImporte * -1 " _
        & "                         WHEN mc.nMovImporte > 0 THEN 0 END, 0) ) as nHaber, " _
        & "       cls.cCtaCaracter, '2' " _
        & "FROM   MovCta mc JOIN Mov M ON M.nMovNro = mc.nMovNro JOIN CtaCont c ON c.cCtaContCod = mc.cCtaContCod JOIN " & sCentralCom & "CtaContClase cls ON mc.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' " _
        & "WHERE M.nMovEstado = '" & gMovEstContabMovContable & "' and not M.nMovFlag in ('" & gMovFlagEliminado & "','" & gMovFlagModificado & "') and SubString(m.cMovNro,1,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' " _
        & "      " & sCondBala & " and mc.cCtaContCod IS NOT NULL and mc.cCtaContCod <> '' " & IIf(pbBCient Or pbConCierreAnual, "", " and not M.cOpeCod LIKE '" & Left(gContCierreAnual, 5) & "%' ") _
        & "GROUP BY " & sCta & ", cls.cCtaCaracter "
   dbConec.Ejecutar psSql

Exit Sub
InsertaMovimientosMesErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaMovimientosMes Method")
End Sub

Public Sub MayorizacionBalance(pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer, Optional pbConCierreAnual As Boolean = False, Optional psCodCmac = "112")
Dim sCondBala As String
Dim sCta      As String
On Error GoTo InsertaMovimientosMesErr
If psCodCmac = "112" Then
  psSql = "INSERT BalanceEstad " _
        & "SELECT '" & pnTipoBala & "','" & pnMoneda & "', '" & Format(pnAnio, "0000") & "','" _
        & Format(pnMes + IIf(pbConCierreAnual, 1, 0), "00") & "', c.cCtaContCod, c.cCtaContDesc, SUM(z.nSaldoIniImporte) as nSaldoIni, " _
        & "       SUM( z.nDebe ) as nDebe, SUM( z.nHaber) as nHaber, " _
        & "       ISNULL(SUM(CASE WHEN cCtaClasifica = 'D' THEN z.nSaldoIniImporte + z.nDebe - z.nHaber " _
        & "                       WHEN cCtaClasifica = 'A' THEN z.nSaldoIniImporte - z.nDebe + z.nHaber END) ,0) as nSaldoFinImporte " _
        & "FROM   BalanceTemp as z JOIN " & sCentralCom & "CtaCont c ON z.cCtaContCod LIKE RTRIM(c.cCtaContCod) + '%' " _
        & "WHERE  cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda & "' " _
        & "GROUP BY c.cCtaContCod, c.cCtaContDesc " _
        & "ORDER BY c.cCtaContCod "
Else
  psSql = "INSERT BalanceEstad " _
        & "SELECT '" & pnTipoBala & "','" & pnMoneda & "', '" & Format(pnAnio, "0000") & "','" _
        & Format(pnMes + IIf(pbConCierreAnual, 1, 0), "00") & "', c.cCtaContCod, c.cCtaContDesc, CASE WHEN c.cCtaContCod LIKE '[678]' THEN SUM(z.nSaldoIniImporte * CASE WHEN cCtaClasifica = 'D' THEN 1 ELSE -1 END) ELSE SUM(z.nSaldoIniImporte) END as nSaldoIni, " _
        & "      SUM( z.nDebe) as nDebe, SUM( z.nHaber ) as nHaber, " _
        & "      CASE WHEN c.cCtaContCod LIKE '[678]' THEN " _
        & "             ISNULL(SUM(CASE WHEN cCtaClasifica = 'D' THEN z.nSaldoIniImporte + z.nDebe - z.nHaber " _
        & "                             WHEN cCtaClasifica = 'A' THEN z.nSaldoIniImporte - z.nDebe + z.nHaber END * CASE WHEN cCtaClasifica = 'D' THEN 1 ELSE -1 END ) ,0 ) " _
        & "        ELSE ISNULL(SUM(CASE WHEN cCtaClasifica = 'D' THEN z.nSaldoIniImporte + z.nDebe - z.nHaber " _
        & "                             WHEN cCtaClasifica = 'A' THEN z.nSaldoIniImporte - z.nDebe + z.nHaber END ) ,0 ) END as nSaldoFinImporte " _
        & "FROM   BalanceTemp as z JOIN " & sCentralCom & "CtaCont c ON z.cCtaContCod LIKE RTRIM(c.cCtaContCod) + '%' " _
        & "WHERE  cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda & "' " _
        & "GROUP BY c.cCtaContCod, c.cCtaContDesc " _
        & "ORDER BY c.cCtaContCod "
End If
   dbConec.CommandTimeout 3000
   dbConec.Ejecutar psSql
Exit Sub
InsertaMovimientosMesErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaMovimientosMes Method")
End Sub

Public Function EjecutaBatch() As Integer
On Error GoTo EjecutaBatchErr
    EjecutaBatch = dbConec.EjecutarBatch
    Exit Function
EjecutaBatchErr:
    Call RaiseError(MyUnhandledError, "DBalanceCont:EjecutaBatch Method")
End Function

Public Sub EliminaUtilidadAcumulada(pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer, Optional lbEjecutarBatch As Boolean = False)
On Error GoTo EliminaUtilidadAcumuladaErr
   psSql = "DELETE UtilidadAcumula WHERE cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
        & "' and cUtilidadMes = '" & Format(pnMes, "00") & "' and " _
        & "cUtilidadAnio = '" & Format(pnAnio, "0000") & "'"
   If lbEjecutarBatch Then
      dbConec.AdicionaCmdBatch psSql
   Else
      dbConec.Ejecutar psSql
   End If
Exit Sub
EliminaUtilidadAcumuladaErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:EliminaUtilidadAcumulada Method")
End Sub
   
Public Sub InsertaUtilidadAcumulada(pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer, pnUtilidad As Currency, Optional lbEjecutarBatch As Boolean = False)
On Error GoTo InsertaUtilidadAcumuladaErr
   psSql = "INSERT UtilidadAcumula VALUES ('" & pnTipoBala & "', '" & pnMoneda & "', '" & Format(pnAnio, "0000") & "', '" & Format(pnMes, "00") & "', " & pnUtilidad & ")"
   If lbEjecutarBatch Then
      dbConec.AdicionaCmdBatch psSql
   Else
      dbConec.Ejecutar psSql
   End If
Exit Sub
InsertaUtilidadAcumuladaErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaUtilidadAcumulada Method")
End Sub

Public Sub InsertaBalanceTmpSaldos(pnTipoBala As Integer, pnMoneda As Integer, psFecha As String, Optional lbConsolCta As Boolean = False)
Dim sCampoCta As String
On Error GoTo InsertaBalanceTmpSaldosErr
   If pnMoneda = 0 And lbConsolCta Then
      sCampoCta = "substring(Cta.cCtaContCod,1,2) + '0' + substring(Cta.cCtaContCod,4,20)"
      psSql = " INSERT Into BalanceTemp " _
         & " SELECT '" & pnTipoBala & "','" & pnMoneda & "', " & sCampoCta & ", SUM(Cta.nCtaSaldoImporte), 0,0,cls.cCtaCaracter,'1'  " _
         & " FROM   CtaSaldo Cta JOIN " & sCentralCom & "CtaContClase cls ON cta.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' " _
         & " WHERE  cta.nCtaSaldoImporte <> 0 and Cta.dCtaSaldofecha in (SELECT MAX(Cta2.dCtaSaldofecha) From Ctasaldo Cta2 where cta2.cCtaContCod = Cta.cCtacontCod and Cta2.dCtaSaldofecha <= '" & psFecha & "') " & IIf(pnMoneda = 0, "", "and substring(cCtaContCod,3,1) = '" & pnMoneda & "'") _
         & " GROUP BY " & sCampoCta & ", cls.cCtaCaracter"
   Else
      sCampoCta = "Cta.cCtaContCod"
      psSql = " INSERT Into BalanceTemp " _
         & " SELECT '" & pnTipoBala & "','" & pnMoneda & "', " & sCampoCta & ", Cta.nCtaSaldoImporte, 0,0,cls.cCtaCaracter,'1'  " _
         & " FROM   CtaSaldo Cta JOIN " & sCentralCom & "CtaContClase cls ON cta.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' " _
         & " WHERE  cta.nCtaSaldoImporte <> 0 and Cta.dCtaSaldofecha in (SELECT MAX(Cta2.dCtaSaldofecha) From Ctasaldo Cta2 where cta2.cCtaContCod = Cta.cCtacontCod and Cta2.dCtaSaldofecha <= '" & psFecha & "') " & IIf(pnMoneda = 0, "", "and substring(Cta.cCtaContCod,3,1) = '" & pnMoneda & "'")
   End If
         
   dbConec.Ejecutar psSql
Exit Sub
InsertaBalanceTmpSaldosErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceTmpSaldos Method")
End Sub

Public Sub InsertaBalanceGen(psCodigo As Integer, pnMN As Currency, pnME As Currency, pnTotal As Currency, pnTotAj As Currency, psFecha As String, psBalanceCate As String, Optional lbEjecutarBatch As Boolean = False)
On Error GoTo InsertaBalanceGenErr
   psSql = "INSERT INTO BalanceGen(cCodigo,nMN,nME,nTotal,nTotAj,cFecha,cBalanceCate) VALUES('" _
           & psCodigo & "'," & pnMN & "," & pnME & "," & pnTotal & "," & pnTotAj & ",'" & psFecha & "','" & psBalanceCate & "')"
   If lbEjecutarBatch Then
      dbConec.AdicionaCmdBatch psSql
   Else
      dbConec.Ejecutar psSql
   End If
Exit Sub
InsertaBalanceGenErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceGen Method")
End Sub

Public Function CargaBCient(psFecha As String, pnNroDigitos As Integer, pbConCaluloUtilidad As Boolean, psCMAC As String) As Recordset
    Dim lsSub As String
   
   If pbConCaluloUtilidad = False Then
      psSql = "SELECT * from BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & "LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " Order By SubString(cCtaContCod,1,11) + SubString(cCtaContCod,13,17) + cMoneda"
   Else
      
      psSql = " SELECT dFecha, cCtaContCod, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte, cMoneda, cCtaContCodAnt from BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt <> '3'"
      'psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6100' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6101' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5') And cMoneda = '0' " _
            & " Group By cMoneda"
      'psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6101' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6101' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5') And cMoneda = '0' " _
            & " Group By cMoneda"
      'psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6600' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','62','64') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6601' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','62','63','64','65') And cMoneda = '0' " _
            & " Group By cMoneda"
      'psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6601' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6601' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501') And cMoneda = '0' " _
            & " Group By cMoneda"
      'psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6901' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6901' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501','6701','6801') And cMoneda = '0' " _
            & " Group By cMoneda"
      'psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '69' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401','6902') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '69' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501','6701','6801','6962') And cMoneda = '0' " _
            & " Group By cMoneda"
      'psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nDebe Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nHaber Else nHaber end) nHaber, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','62','63','64','65','67','68','69') And cMoneda = '0' " _
            & " Group By cMoneda"
      'psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3010' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('30','3060') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('30','3060','3030') And cMoneda = '0' " _
            & " Group By cMoneda"
'ANT      'psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nDebe Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nHaber Else nHaber end) nHaber, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','62','63','64','65','67','68','69') And cMoneda in ('0','6') " _
            & " Group By cMoneda"
        
        lsSub = " (SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '30' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte, cMoneda, '30' + cMoneda + '0' cCtaContCodAnt " _
              & "  From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 And " _
              & "  LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('3') And cMoneda in ('0','1','6')) As AAA " _
              & "  Inner Join (SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
              & "   From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
              & "   LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
              & "   Group By cMoneda " _
              & "   Union " _
              & "   SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '391' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391' cCtaContCodAnt " _
              & "   From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
              & "   LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2')) As BBB On AAA.cMoneda = BBB.cMoneda  "
        
        lsSub = " Select AAA.dFecha, AAA.cCtaContCod, AAA.nSaldoIniImporte + BBB.nSaldoIniImporte nSaldoIniImporte , AAA.nDebe + BBB.nHaber nDebe , AAA.nHaber + BBB.nDebe  nHaber, AAA.nSaldoFinImporte + BBB.nSaldoFinImporte nSaldoFinImporte, AAA.cMoneda, AAA.cCtaContCodAnt From " & lsSub
        
        psSql = psSql & " Union " & lsSub
        
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
'            & " Group By cMoneda"
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '000000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '1000000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
'            & " Group By cMoneda"
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '100000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '3911' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '1010000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '101' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
'            & " Group By cMoneda"
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '101000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391101' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "

'*********************1
'***********************
        psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
            & " Group By cMoneda"

        psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '391' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "

        psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '39' + cMoneda + '' + '1000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
            & " Group By cMoneda"

        psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '391' + '100000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '3911' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "

        psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '39' + cMoneda + '' + '1010000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '101' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
            & " Group By cMoneda"

        psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '391' + '101000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391101' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "

'***********************


''***********************
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
'            & " Group By cMoneda"
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '1000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
'            & " Group By cMoneda"
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '100000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '3911' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '1010000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '101' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
'            & " Group By cMoneda"
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '101000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391101' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
'
''***********************

'      psSql = psSql & " Union " _
            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '1' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nDebe Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nHaber Else nHaber end) nHaber, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','62','63','64','65','67','68','69') And cMoneda in ('0','1','2','6') " _
            & " Group By cMoneda"
'      psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3911' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nDebe * -1  Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nHaber Else nHaber * -1 end) nHaber, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '3911' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('401','501','401','501','6211','6311','6411','6511','6211','6311','6411','6511','6711','6811','6711','6811') And cMoneda In ('1') " _
'            & " Group By cMoneda"
'      psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3921' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nDebe * -1 Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nHaber Else nHaber * -1 end) nHaber, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '3921' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('402','502','402','502','6221','6321','6421','6521','6221','6321','6421','6521','6721','6821','6721','6821') And cMoneda In ('2') " _
'            & " Group By cMoneda"
'      psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3961' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('506','506','6261','6461','6261','6461','3961') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('506','506','6261','6461','6261','6461','3961') then nDebe * -1 Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('506','506','6261','6461','6261','6461','3961') then nHaber Else nHaber * -1 end) nHaber, sum(Case when cCtaContCodAnt In ('506','506','6261','6461','6261','6461','3961') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '3961' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('406','506','406','506','6261','6361','6461','6561','6261','6361','6461','6561','6761','6861','6761','6861','3961') And cMoneda In ('6') " _
'            & " Group By cMoneda"

'      psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3902' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401','6902') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501','6701','6801','6902','401','501','402','502','6211','6311','6411','6511','6221','6321','6421','6521','6711','6811','6721','6821','3961') And cMoneda In ('0','1','2','6') " _
'            & " Group By cMoneda"
'      psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '390201' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401','6902') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501','6701','6801','6902','401','501','402','502','6211','6311','6411','6511','6221','6321','6421','6521','6711','6811','6721','6821','3961') And cMoneda In ('0','1','2','6') " _
'            & " Group By cMoneda"
'      psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3962100' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('3962') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('3962') And cMoneda In ('1','2','6') " _
'            & " Group By cMoneda"
            
      psSql = " Select * From (" & psSql & " ) As AAA Order By SubString(cCtaContCod,1,11) + SubString(cCtaContCod,13,17) + cMoneda"
   
   End If
   
   Set CargaBCient = dbConec.CargaRecordSet(psSql)
   
End Function

Public Sub EliminaBCient(pdfecha As Date, Optional psCodigo As String = "")
On Error GoTo EliminaBCientErr
   psSql = "DELETE BCient " _
         & "WHERE dFecha = '" & Format(pdfecha, gsFormatoFecha) & "' " & IIf(psCodigo = "", "", " and cCtaContCod = '" & psCodigo & "'")
   dbConec.Ejecutar psSql
Exit Sub
EliminaBCientErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:EliminaBCient Method")
End Sub

Public Sub InsertaBCient(pdfecha As Date, psCodigo, pnCateBala As Integer, pnMoneda As Integer)
On Error GoTo InsertaBCientErr
  psSql = "INSERT BCient " _
        & "SELECT '" & Format(pdfecha, gsFormatoFecha) & "', SUBSTRING('" & psCodigo & "' + c.cCtaContCod + '" & String(20, "0") & "',1,29), SUM(z.nSaldoIniImporte) as nSaldoIni, " _
        & "       SUM( z.nDebe ) as nDebe, SUM( z.nHaber) as nHaber, " _
        & "       ISNULL(SUM(CASE WHEN cCtaClasifica = 'D' THEN z.nSaldoIniImporte + z.nDebe - z.nHaber " _
        & "                       WHEN cCtaClasifica = 'A' THEN z.nSaldoIniImporte - z.nDebe + z.nHaber END) ,0) as nSaldoFinImporte, cBalanceTipo, c.cCtaContcod " _
        & "FROM   BalanceTemp as z JOIN " & sCentralCom & "CtaCont c ON z.cCtaContCod LIKE c.cCtaContCod + '%' " _
        & "WHERE  cBalanceCate = '" & pnCateBala & "' and cBalanceTipo = '" & pnMoneda & "' " _
        & "GROUP BY c.cCtaContCod, c.cCtaContDesc, cBalanceTipo " _
        & "ORDER BY c.cCtaContCod "
  dbConec.CommandTimeout 500
  dbConec.Ejecutar psSql
Exit Sub
InsertaBCientErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBCient Method")
End Sub

Public Sub ActualizaBCient(pdfecha As Date, psCodigo, pnCateBala As Integer, pnMoneda As Integer)
On Error GoTo ActualizaBCientErr
   psSql = "UPDATE BCient Set cCtaContCod = '" & psCodigo & "' + cCtaContCodAnt + '0" & pnMoneda & "00000000000000000' " _
       & " WHERE datediff(dd,dFecha,'" & Format(pdfecha, gsFormatoFecha) & "') = 0 and " _
       & "       LEN(ltrim(rtrim(cCtacontcodant))) = 1 and cMoneda = '" & pnMoneda & "'"
   dbConec.Ejecutar psSql
                                              
   psSql = "UPDATE BCient Set cCtaContCod = '" & psCodigo & "' + cCtaContCodAnt + '" & pnMoneda & "00000000000000000' " _
       & " WHERE datediff(dd,dFecha,'" & Format(pdfecha, gsFormatoFecha) & "') = 0 and " _
       & " LEN(ltrim(rtrim(cCtacontcodant))) = 2 and cMoneda = '" & pnMoneda & "'"
   dbConec.Ejecutar psSql
Exit Sub
ActualizaBCientErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ActualizaBCient Method")
End Sub


Public Sub EliminaBalanceSectorial(pdfecha As Date, Optional psCodigo As String = "", Optional lbEjecutarBatch As Boolean = False)
On Error GoTo EliminaBalanceErr
   psSql = "DELETE BalanceSectorial " _
         & "WHERE nBalanceMes = " & Month(pdfecha) & " and nBalanceAnio = " & Year(pdfecha) & IIf(psCodigo = "", "", " and cCodigo = '" & psCodigo & "'")
   If lbEjecutarBatch Then
      dbConec.AdicionaCmdBatch psSql
   Else
      dbConec.Ejecutar psSql
   End If
Exit Sub
EliminaBalanceErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:EliminaBalanceSectorial Method")
End Sub

Public Sub InsertaBalanceSectorial(psCodigo As String, pnMN As Currency, pnME As Currency, pnMNAj As Currency, psFecha As String, Optional lbEjecutarBatch As Boolean = False)
On Error GoTo InsertaBalanceSectorialErr
   psSql = "INSERT INTO BalanceSectorial(cCodigo,nBalanceAnio, nBalanceMes, nMNAj, nMNHist, nME) VALUES('" _
           & psCodigo & "'," & Year(CDate(psFecha)) & "," & Month(CDate(psFecha)) & "," & pnMNAj & "," & pnMN & "," & pnME & ")"
   If lbEjecutarBatch Then
      dbConec.AdicionaCmdBatch psSql
   Else
      dbConec.Ejecutar psSql
   End If
Exit Sub
InsertaBalanceSectorialErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Sub

Public Sub ActualizaBalanceSectorial(psCodigo As String, pnMNHist As Currency, pnME As Currency, pnMNAj As Currency, psFecha As String, Optional lbEjecutarBatch As Boolean = False)
On Error GoTo ActualizaBalanceSectorialErr
   psSql = "UPDATE BalanceSectorial SET nMNAj = " & pnMNAj & ", nMNHist = " & pnMNHist & ", nME = " & pnME & " " _
           & "WHERE cCodigo = '" & psCodigo & "' and nBalanceAnio = " & Year(CDate(psFecha)) & " and nBalanceMes = " & Month(CDate(psFecha))
   If lbEjecutarBatch Then
      dbConec.AdicionaCmdBatch psSql
   Else
      dbConec.Ejecutar psSql
   End If
Exit Sub
ActualizaBalanceSectorialErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ActualizaBalanceSectorial Method")
End Sub



