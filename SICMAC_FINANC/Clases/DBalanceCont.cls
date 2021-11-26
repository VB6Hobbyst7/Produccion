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
'dbConec.CommandTimeOut 4500
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
'    psSql = "DELETE BalanceTemp " _
'        & "WHERE cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda & "'"
   
    psSql = "exec stp_del_BalanceTemp '" & pnTipoBala & "','" & pnMoneda & "'" 'PEAC 20210420
   
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

psSql = "Select nMovNro FROM Mov WHERE cMovNro LIKE '" & Format(lsFecha, gsFormatoMovFecha) & "%' and cOpeCod liKE '70185%' and nmovflag = 0 "
Set prs = dbConec.CargaRecordSet(psSql)
If Not prs.EOF And lnMes = 12 Then
   lnMes = 13
End If
'ALPA 20090510*************************************************
'psSql = "select ISNULL(b1.cCtaContCod,b2.cCtaContCod), nSaldoIniImporte, nSaldoFinImporte " _
'      & " From ( Select cs.cCtaContCod, nCtaSaldoImporte nSaldoIniImporte " _
'      & "        from CtaSaldo cs, CtaCont cc " _
'      & " where dCtaSaldoFecha = (Select Max(dCtaSaldoFecha) FROM CtaSaldo cs1 WHERE cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha < = '" & lsFecha & "') " _
'      & "              and cc.cCtaContCod like cs.cCtaContCod + '%' " _
'      & "        group by cs.cCtaContCod, nCtaSaldoImporte " _
'      & "        having count(*) = 1 " _
'      & "      ) b2 LEFT JOIN " _
'      & "      ( " _
'      & "        Select cCtaContCod, nSaldoFinImporte from BalanceEstad " _
'      & "        where cBalanceCate = 2 and cBalanceTipo in (1,2) and cBalanceMes = '" & Format(lnMes, "00") & "' and cBalanceAnio = '" & Format(lnAnio, "0000") & "' " _
'      & "      ) b1 ON b1.cCtaContCod = b2.cCtaContCod " _
'      & "      Where IsNull(b1.nSaldoFinImporte, 0) <> IsNull(b2.nSaldoIniImporte, 0) "
psSql = "exec stp_sel_ValidaSaldosIniciales '" & Format(lsFecha, "YYYY/MM/DD") & "','" & Format(lnMes, "00") & "','" & Format(lnAnio, "0000") & "'"
'*********************************************************************
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
      sCondBala = "and ( (Substring(a.cCtaContCod,3,1) Like '[13]' or substring(a.cCtaContCod,1,1) IN (" & sCtaMN & " ) " _
                & " or substring(a.cCtaContCod,1,2) IN (" & sCtaMN & " ) ) " _
                & IIf(pnTipoBala = 2, " or  substring(a.cCtaContCod,3,1) = '6' )", " and substring(a.cCtaContCod,3,1) <> '6' )")
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
   'dbConec.CommandTimeOut 3000
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

'sCtaMN = "SELECT cCtaContCod FROM " & sCentralCom & "CtaContClase WHERE nCtaMoneda = 1"
'If pbBCient Then
'   If pnMoneda = 0 Then
'      sCta = "substring(mc.cCtaContCod,1,2) + '0' + substring(mc.cCtaContCod,4,20)"
'      sCondBala = ""
'   ElseIf pnMoneda = 3 And pnTipoBala <> 5 Then
'      sCta = "mc.cCtaContCod"
'      sCondBala = " "
'   Else
'      sCondBala = " and (SubString(mc.cCtaContCod,3,1) = '" & pnMoneda & "') "
'      sCta = "mc.cCtaContCod"
'   End If
'Else
'   If pnMoneda <> 0 And pnMoneda <> 9 Then
'      If pnMoneda = 1 Then
'         sCondBala = "and ( (Substring(mc.cCtaContCod,3,1) Like '[13]' or substring(mc.cCtaContCod,1,1) IN (" & sCtaMN & ") " _
'                   & " or substring(mc.cCtaContCod,1,2) IN (" & sCtaMN & ") ) " _
'                   & IIf(pnTipoBala = 2, "or  substring(mc.cCtaContCod,3,1) = '6' )", " and substring(mc.cCtaContCod,3,1) <> '6' )")
'      ElseIf pnMoneda = 2 Then
'         sCondBala = "and ( (Substring(mc.cCtaContCod,3,1) = '2' and NOT substring(mc.cCtaContCod,1,1) IN (" & sCtaMN & ") " _
'                   & " and NOT substring(mc.cCtaContCod,1,2) IN (" & sCtaMN & ") )" _
'                   & " and substring(mc.cCtaContCod,3,1) <> '6' )"
'      Else
'         sCondBala = "and ( Substring(mc.cCtaContCod,3,1) = '" & pnMoneda & "' " _
'                   & IIf(pnTipoBala = 2, "or  substring(mc.cCtaContCod,3,1) = '6' )", " )")
'      End If
'      sCta = "mc.cCtaContCod"
'   ElseIf pnMoneda = 9 Then
'      sCta = "mc.cCtaContCod"
'      sCondBala = " and (Substring(mc.cCtaContCod,3,1) IN ('1','2','3','4','5') " & IIf(pnTipoBala = 2, " or substring(mc.cCtaContCod,3,1) = '6' )", ") ")
'   Else
'      sCta = "SubString(mc.cCtaContCod,1,2) + '0' + SubString(mc.cCtaContCod,4,20) "
'      sCondBala = " and (Substring(mc.cCtaContCod,3,1) IN ('1','2','3','4','5') " & IIf(pnTipoBala = 2, " or substring(mc.cCtaContCod,3,1) = '6' )", ") ")
'   End If
'End If
'   psSql = "INSERT BALANCETEMP " _
'        & "SELECT '" & pnTipoBala & "','" & pnMoneda & "'," & sCta & ", 0," _
'        & "       SUM( ISNULL( CASE WHEN mc.nMovImporte > 0 THEN mc.nMovImporte " _
'        & "                         WHEN mc.nMovImporte < 0 THEN 0 END, 0) ) as nDebe, " _
'        & "       SUM( ISNULL( CASE WHEN mc.nMovImporte < 0 THEN mc.nMovImporte * -1 " _
'        & "                         WHEN mc.nMovImporte > 0 THEN 0 END, 0) ) as nHaber, " _
'        & "       cls.cCtaCaracter, '2' " _
'        & "FROM   MovCta mc JOIN Mov M ON M.nMovNro = mc.nMovNro JOIN CtaCont c ON c.cCtaContCod = mc.cCtaContCod JOIN " & sCentralCom & "CtaContClase cls ON mc.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' " _
'        & "WHERE M.nMovEstado = '" & gMovEstContabMovContable & "' and not M.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & "," & gMovFlagExtornado & ") and SubString(m.cMovNro,1,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' " _
'        & "      " & sCondBala & " and mc.cCtaContCod IS NOT NULL and mc.cCtaContCod <> '' " & IIf(pbBCient Or pbConCierreAnual, "", " and not M.cOpeCod LIKE '" & Left(gContCierreAnual, 5) & "%' ") _
'        & "GROUP BY " & sCta & ", cls.cCtaCaracter "

    psSql = "exec stp_ins_MovimientosMes '" & pnTipoBala & "','" & pnMoneda & "','" & psFechaDel & "','" & psFechaAl & "'," & IIf(pbBCient, 1, 0) & "," & IIf(pbConCierreAnual, 1, 0) & ",'" & gContCierreAnual & "' " 'PEAC 20210420
   
   dbConec.Ejecutar psSql

Exit Sub
InsertaMovimientosMesErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaMovimientosMes Method")
End Sub

Public Sub CalculaSaldo_39_69_Historico()
Dim sCondBala As String
Dim sCta      As String
Dim sCtaMN    As String
On Error GoTo CalculaSaldo_39_69_HistoricoErr
   
   psSql = "spBCient_Calcula_Utilidad "
   dbConec.Ejecutar psSql

Exit Sub
CalculaSaldo_39_69_HistoricoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:CalculaSaldo_39_69_Historico Method")
End Sub

Public Sub MayorizacionBalance(pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer, Optional pbConCierreAnual As Boolean = False, Optional psCodCmac = "112")
Dim sCondBala As String
Dim sCta      As String
On Error GoTo InsertaMovimientosMesErr
If psCodCmac = "112" Then
  'Formato mayorizacion normal para la cuenta 6
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
  'Formato mayorizacion depende del CtaContClase
'  psSql = "INSERT BalanceEstad " _
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
  
  
  psSql = "INSERT BalanceEstad " _
        & "SELECT '" & pnTipoBala & "','" & pnMoneda & "', '" & Format(pnAnio, "0000") & "','" _
        & Format(pnMes + IIf(pbConCierreAnual, 1, 0), "00") & "', c.cCtaContCod, c.cCtaContDesc, CASE WHEN c.cCtaContCod LIKE '[]' THEN SUM(z.nSaldoIniImporte * CASE WHEN cCtaClasifica = 'D' THEN 1 ELSE -1 END) ELSE SUM(z.nSaldoIniImporte) END as nSaldoIni, " _
        & "      SUM( z.nDebe) as nDebe, SUM( z.nHaber ) as nHaber, " _
        & "      CASE WHEN c.cCtaContCod LIKE '[]' THEN " _
        & "             ISNULL(SUM(CASE WHEN cCtaClasifica = 'D' THEN z.nSaldoIniImporte + z.nDebe - z.nHaber " _
        & "                             WHEN cCtaClasifica = 'A' THEN z.nSaldoIniImporte - z.nDebe + z.nHaber END * CASE WHEN cCtaClasifica = 'D' THEN 1 ELSE -1 END ) ,0 ) " _
        & "        ELSE ISNULL(SUM(CASE WHEN cCtaClasifica = 'D' THEN z.nSaldoIniImporte + z.nDebe - z.nHaber " _
        & "                             WHEN cCtaClasifica = 'A' THEN z.nSaldoIniImporte - z.nDebe + z.nHaber END ) ,0 ) END as nSaldoFinImporte " _
        & "FROM   BalanceTemp as z JOIN " & sCentralCom & "CtaCont c ON z.cCtaContCod LIKE RTRIM(c.cCtaContCod) + '%' " _
        & "WHERE  cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda & "' " _
        & "GROUP BY c.cCtaContCod, c.cCtaContDesc " _
        & "ORDER BY c.cCtaContCod "
End If
   dbConec.CommadTimeOut = 7200
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
'EJVG20130423 Se agregó validación en CtaCont
'   If pnMoneda = 0 And lbConsolCta Then
'      sCampoCta = "substring(Cta.cCtaContCod,1,2) + '0' + substring(Cta.cCtaContCod,4,20)"
'      psSql = " INSERT Into BalanceTemp " _
'         & " SELECT '" & pnTipoBala & "','" & pnMoneda & "', " & sCampoCta & ", SUM(Cta.nCtaSaldoImporte), 0,0,cls.cCtaCaracter,'1'  " _
'         & " FROM   CtaSaldo Cta JOIN " & sCentralCom & "CtaContClase cls ON cta.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' " _
'         & " INNER JOIN CtaCont CC ON CC.cCtaContCod = Cta.cCtaContCod " _
'         & " WHERE  cta.nCtaSaldoImporte <> 0 and Cta.dCtaSaldofecha in (SELECT MAX(Cta2.dCtaSaldofecha) From Ctasaldo Cta2 where cta2.cCtaContCod = Cta.cCtacontCod and Cta2.dCtaSaldofecha <= '" & psFecha & "') " & IIf(pnMoneda = 0, "", "and ( substring(cCtaContCod,3,1) = '" & pnMoneda & "'" & ")") _
'         & " GROUP BY " & sCampoCta & ", cls.cCtaCaracter"
'   Else
'      sCampoCta = "Cta.cCtaContCod"
'      psSql = " INSERT Into BalanceTemp " _
'         & " SELECT '" & pnTipoBala & "','" & pnMoneda & "', " & sCampoCta & ", Cta.nCtaSaldoImporte, 0,0,cls.cCtaCaracter,'1'  " _
'         & " FROM   CtaSaldo Cta JOIN " & sCentralCom & "CtaContClase cls ON cta.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' " _
'         & " INNER JOIN CtaCont CC ON CC.cCtaContCod = Cta.cCtaContCod " _
'         & " WHERE  cta.nCtaSaldoImporte <> 0 and Cta.dCtaSaldofecha in (SELECT MAX(Cta2.dCtaSaldofecha) From Ctasaldo Cta2 where cta2.cCtaContCod = Cta.cCtacontCod and Cta2.dCtaSaldofecha <= '" & psFecha & "') " & IIf(pnMoneda = 0, "", "and ( substring(Cta.cCtaContCod,3,1) = '" & pnMoneda & "'" & ")")
'   End If
         
    psSql = "exec stp_ins_BalanceTemp '" & pnTipoBala & "','" & pnMoneda & "','" & Format(psFecha, "yyyymmdd") & "', " & IIf(lbConsolCta, 1, 0) 'PEAC 20210420
         
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

'Public Function CargaBCient(psFecha As String, pnNroDigitos As Integer, pbConCaluloUtilidad As Boolean, psCMAC As String) As Recordset
'    Dim lsSub As String
'
'   If pbConCaluloUtilidad = False Then
'      psSql = "SELECT * from BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & "LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " Order By SubString(cCtaContCod,1,11) + SubString(cCtaContCod,13,17) + cMoneda"
'   Else
'
'      psSql = " SELECT dFecha, cCtaContCod, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte, cMoneda, cCtaContCodAnt from BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt <> '3'"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6100' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6101' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5') And cMoneda = '0' " _
'            & " Group By cMoneda"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6101' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6101' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5') And cMoneda = '0' " _
'            & " Group By cMoneda"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6600' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','62','64') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6601' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','62','63','64','65') And cMoneda = '0' " _
'            & " Group By cMoneda"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6601' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6601' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501') And cMoneda = '0' " _
'            & " Group By cMoneda"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6901' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6901' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501','6701','6801') And cMoneda = '0' " _
'            & " Group By cMoneda"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '69' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401','6902') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '69' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501','6701','6801','6962') And cMoneda = '0' " _
'            & " Group By cMoneda"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nDebe Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nHaber Else nHaber end) nHaber, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','62','63','64','65','67','68','69') And cMoneda = '0' " _
'            & " Group By cMoneda"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3010' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('30','3060') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('30','3060','3030') And cMoneda = '0' " _
'            & " Group By cMoneda"
''ANT      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nDebe Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nHaber Else nHaber end) nHaber, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','62','63','64','65','67','68','69') And cMoneda in ('0','6') " _
'            & " Group By cMoneda"
'
'        lsSub = " (SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '30' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte, cMoneda, '30' + cMoneda + '0' cCtaContCodAnt " _
'              & "  From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 And " _
'              & "  LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('3') And cMoneda in ('0','1','6')) As AAA " _
'              & "  Inner Join (SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
'              & "   From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'              & "   LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
'              & "   Group By cMoneda " _
'              & "   Union " _
'              & "   SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '391' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391' cCtaContCodAnt " _
'              & "   From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'              & "   LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2')) As BBB On AAA.cMoneda = BBB.cMoneda  "
'
'        lsSub = " Select AAA.dFecha, AAA.cCtaContCod, AAA.nSaldoIniImporte + BBB.nSaldoIniImporte nSaldoIniImporte , AAA.nDebe + BBB.nHaber nDebe , AAA.nHaber + BBB.nDebe  nHaber, AAA.nSaldoFinImporte + BBB.nSaldoFinImporte nSaldoFinImporte, AAA.cMoneda, AAA.cCtaContCodAnt From " & lsSub
'
'        psSql = psSql & " Union " & lsSub
'
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
''            & " Group By cMoneda"
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '000000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '1000000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
''            & " Group By cMoneda"
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '100000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '3911' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '1010000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '101' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
''            & " Group By cMoneda"
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '101000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391101' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
'
''*********************1
''***********************
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
'            & " Group By cMoneda"
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '391' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '39' + cMoneda + '' + '1000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
'            & " Group By cMoneda"
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '391' + '100000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '3911' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '39' + cMoneda + '' + '1010000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '101' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
'            & " Group By cMoneda"
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '391' + '101000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391101' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
'
''***********************
'
'
'''***********************
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
''            & " Group By cMoneda"
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '1000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
''            & " Group By cMoneda"
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '100000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '3911' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '1010000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '101' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
''            & " Group By cMoneda"
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '101000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391101' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
''
'''***********************
'
''      psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '1' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nDebe Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nHaber Else nHaber end) nHaber, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','62','63','64','65','67','68','69') And cMoneda in ('0','1','2','6') " _
'            & " Group By cMoneda"
''      psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3911' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nDebe * -1  Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nHaber Else nHaber * -1 end) nHaber, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '3911' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('401','501','401','501','6211','6311','6411','6511','6211','6311','6411','6511','6711','6811','6711','6811') And cMoneda In ('1') " _
''            & " Group By cMoneda"
''      psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3921' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nDebe * -1 Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nHaber Else nHaber * -1 end) nHaber, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '3921' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('402','502','402','502','6221','6321','6421','6521','6221','6321','6421','6521','6721','6821','6721','6821') And cMoneda In ('2') " _
''            & " Group By cMoneda"
''      psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3961' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('506','506','6261','6461','6261','6461','3961') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('506','506','6261','6461','6261','6461','3961') then nDebe * -1 Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('506','506','6261','6461','6261','6461','3961') then nHaber Else nHaber * -1 end) nHaber, sum(Case when cCtaContCodAnt In ('506','506','6261','6461','6261','6461','3961') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '3961' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('406','506','406','506','6261','6361','6461','6561','6261','6361','6461','6561','6761','6861','6761','6861','3961') And cMoneda In ('6') " _
''            & " Group By cMoneda"
'
''      psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3902' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401','6902') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501','6701','6801','6902','401','501','402','502','6211','6311','6411','6511','6221','6321','6421','6521','6711','6811','6721','6821','3961') And cMoneda In ('0','1','2','6') " _
''            & " Group By cMoneda"
''      psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '390201' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401','6902') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501','6701','6801','6902','401','501','402','502','6211','6311','6411','6511','6221','6321','6421','6521','6711','6811','6721','6821','3961') And cMoneda In ('0','1','2','6') " _
''            & " Group By cMoneda"
''      psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3962100' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('3962') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('3962') And cMoneda In ('1','2','6') " _
''            & " Group By cMoneda"
'
'      psSql = " Select * From (" & psSql & " ) As AAA Order By SubString(cCtaContCod,1,11) + SubString(cCtaContCod,13,17) + cMoneda"
'
'   End If
'
'   Set CargaBCient = dbConec.CargaRecordSet(psSql)
'
'End Function

Public Function CargaBCient(psFecha As String, pnNroDigitos As Integer, pbConCaluloUtilidad As Boolean, psCMAC As String) As Recordset
    Dim lsSub As String
   
'   If pbConCaluloUtilidad = False Then
'      psSql = "SELECT * from BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & "LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " Order By SubString(cCtaContCod,1,11) + SubString(cCtaContCod,13,17) + cMoneda"
'   Else
'
'      psSql = " SELECT dFecha, cCtaContCod, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte, cMoneda, cCtaContCodAnt from BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt <> '3'"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6100' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6101' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5') And cMoneda = '0' " _
'            & " Group By cMoneda"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6101' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6101' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5') And cMoneda = '0' " _
'            & " Group By cMoneda"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6600' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','62','64') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6601' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','62','63','64','65') And cMoneda = '0' " _
'            & " Group By cMoneda"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6601' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6601' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501') And cMoneda = '0' " _
'            & " Group By cMoneda"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '6901' + '00000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '6901' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501','6701','6801') And cMoneda = '0' " _
'            & " Group By cMoneda"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '69' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401','6902') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '69' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501','6701','6801','6962') And cMoneda = '0' " _
'            & " Group By cMoneda"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nDebe Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nHaber Else nHaber end) nHaber, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','62','63','64','65','67','68','69') And cMoneda = '0' " _
'            & " Group By cMoneda"
'      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3010' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('30','3060') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('30','3060','3030') And cMoneda = '0' " _
'            & " Group By cMoneda"
''ANT      'psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nDebe Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nHaber Else nHaber end) nHaber, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','62','63','64','65','67','68','69') And cMoneda in ('0','6') " _
'            & " Group By cMoneda"
'
'        lsSub = " (SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '30' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte, cMoneda, '30' + cMoneda + '0' cCtaContCodAnt " _
'              & "  From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 And " _
'              & "  LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('3') And cMoneda in ('0','1','6')) As AAA " _
'              & "  Inner Join (SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
'              & "   From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'              & "   LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
'              & "   Group By cMoneda " _
'              & "   Union " _
'              & "   SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '391' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391' cCtaContCodAnt " _
'              & "   From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'              & "   LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2','3')) As BBB On AAA.cMoneda = BBB.cMoneda  "
'
'        lsSub = " Select AAA.dFecha, AAA.cCtaContCod, AAA.nSaldoIniImporte + BBB.nSaldoIniImporte nSaldoIniImporte , AAA.nDebe + BBB.nHaber nDebe , AAA.nHaber + BBB.nDebe  nHaber, AAA.nSaldoFinImporte + BBB.nSaldoFinImporte nSaldoFinImporte, AAA.cMoneda, AAA.cCtaContCodAnt From " & lsSub
'
'        psSql = psSql & " Union " & lsSub
'
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
''            & " Group By cMoneda"
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '000000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '1000000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
''            & " Group By cMoneda"
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '100000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '3911' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '1010000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '101' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
''            & " Group By cMoneda"
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '101000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, sum(nDebe) + sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391101' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
'
''*********************1
''***********************
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
'            & " Group By cMoneda"
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '391' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2','3') "
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '39' + cMoneda + '' + '1000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
'            & " Group By cMoneda"
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '391' + '100000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '3911' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2','3') "
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '39' + cMoneda + '' + '1010000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '101' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
'            & " Group By cMoneda"
'
'        psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "' + '391' + '101000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nHaber) nDebe, sum(nDebe) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391101' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2','3') "
'
''***********************
'
'
'''***********************
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
''            & " Group By cMoneda"
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '1000000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
''            & " Group By cMoneda"
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '100000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '3911' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '' + '1010000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '101' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('0','6') " _
''            & " Group By cMoneda"
''
''        psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '391' + '101000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('1') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(nDebe) nDebe, sum(nHaber) nHaber, sum(Case when cCtaContCodAnt In ('1') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte,  '1' cMoneda, '391101' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('1','2','3') And cMoneda in ('1','2') "
''
'''***********************
'
''      psSql = psSql & " Union " _
'            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '39' + cMoneda + '1' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nDebe Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nHaber Else nHaber end) nHaber, sum(Case when cCtaContCodAnt In ('5','62','64','69') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' + cMoneda + '1' cCtaContCodAnt " _
'            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
'            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','62','63','64','65','67','68','69') And cMoneda in ('0','1','2','6') " _
'            & " Group By cMoneda"
''      psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3911' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nDebe * -1  Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nHaber Else nHaber * -1 end) nHaber, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '3911' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('401','501','401','501','6211','6311','6411','6511','6211','6311','6411','6511','6711','6811','6711','6811') And cMoneda In ('1') " _
''            & " Group By cMoneda"
''      psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3921' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nDebe * -1 Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nHaber Else nHaber * -1 end) nHaber, sum(Case when cCtaContCodAnt In ('501','502','6211','6411','6221','6421') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '3921' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('402','502','402','502','6221','6321','6421','6521','6221','6321','6421','6521','6721','6821','6721','6821') And cMoneda In ('2') " _
''            & " Group By cMoneda"
''      psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3961' + '00000000000000000000',29) cCtaContCod, sum(Case when cCtaContCodAnt In ('506','506','6261','6461','6261','6461','3961') then nSaldoIniImporte Else nSaldoIniImporte * -1 end) nSaldoIniImporte, sum(Case when cCtaContCodAnt In ('506','506','6261','6461','6261','6461','3961') then nDebe * -1 Else nDebe end) nDebe, sum(Case when cCtaContCodAnt In ('506','506','6261','6461','6261','6461','3961') then nHaber Else nHaber * -1 end) nHaber, sum(Case when cCtaContCodAnt In ('506','506','6261','6461','6261','6461','3961') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '3961' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('406','506','406','506','6261','6361','6461','6561','6261','6361','6461','6561','6761','6861','6761','6861','3961') And cMoneda In ('6') " _
''            & " Group By cMoneda"
'
''      psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3902' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401','6902') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501','6701','6801','6902','401','501','402','502','6211','6311','6411','6511','6221','6321','6421','6521','6711','6811','6721','6821','3961') And cMoneda In ('0','1','2','6') " _
''            & " Group By cMoneda"
''      psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '390201' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('5','6201','6401','6902') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('4','5','6201','6301','6401','6501','6701','6801','6902','401','501','402','502','6211','6311','6411','6511','6221','6321','6421','6521','6711','6811','6721','6821','3961') And cMoneda In ('0','1','2','6') " _
''            & " Group By cMoneda"
''      psSql = psSql & " Union " _
''            & " SELECT '" & Format(CDate(psFecha), "yyyy-mm-dd hh:mm:ss") & "' dFecha, left('" & Format(CDate(psFecha), "yyyymm") & psCMAC & "'+ '3962100' + '00000000000000000000',29) cCtaContCod, 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, sum(Case when cCtaContCodAnt In ('3962') then nSaldoFinImporte Else nSaldoFinImporte * -1 end) nSaldoFinImporte, cMoneda, '39' cCtaContCodAnt " _
''            & " From BCIENT Where datediff(dd,dFecha,'" & psFecha & "') = 0 and " _
''            & " LEN(ltrim(rtrim(cCtaContCodAnt)))<=" & pnNroDigitos & " And cCtaContCodAnt in ('3962') And cMoneda In ('1','2','6') " _
''            & " Group By cMoneda"
'
'      psSql = " Select * From (" & psSql & " ) As AAA Order By SubString(cCtaContCod,1,11) + SubString(cCtaContCod,13,17) + cMoneda"
'
'   End If
   
   psSql = "exec stp_sel_BCIENT '" & Format(CDate(psFecha), "yyyymmdd") & "'," & pnNroDigitos & ", " & IIf(pbConCaluloUtilidad, 1, 0) & ",'" & psCMAC & "' " 'PEAC 20210420
   
   Set CargaBCient = dbConec.CargaRecordSet(psSql)
   
End Function



Public Sub EliminaBCient(pdFecha As Date, Optional psCodigo As String = "")
On Error GoTo EliminaBCientErr
'   psSql = "DELETE BCient " _
         & "WHERE dFecha = '" & Format(pdFecha, gsFormatoFecha) & "' " & IIf(psCodigo = "", "", " and cCtaContCod = '" & psCodigo & "'")
   
    psSql = "exec stp_del_Bcient " & Format(pdFecha, "yyyymmdd") & IIf(psCodigo = "", "", psCodigo) 'PEAC 20210420
  
   
   dbConec.CommadTimeOut = 7200
   dbConec.Ejecutar psSql
Exit Sub
EliminaBCientErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:EliminaBCient Method")
End Sub

Public Sub InsertaBCient(pdFecha As Date, psCodigo, pnCateBala As Integer, pnMoneda As Integer)
On Error GoTo InsertaBCientErr
   
'  psSql = "INSERT BCient " _
'        & "SELECT '" & Format(pdFecha, gsFormatoFecha) & "', SUBSTRING('" & psCodigo & "' + c.cCtaContCod + '" & String(20, "0") & "',1,29), SUM(z.nSaldoIniImporte ) as nSaldoIni, " _
'        & "       SUM( z.nDebe ) as nDebe, SUM( z.nHaber) as nHaber, " _
'        & "       ISNULL(SUM(CASE WHEN cCtaClasifica = 'D' THEN z.nSaldoIniImporte + z.nDebe - z.nHaber " _
'        & "                       WHEN cCtaClasifica = 'A' THEN z.nSaldoIniImporte - z.nDebe + z.nHaber END ) ,0) as nSaldoFinImporte, cBalanceTipo, c.cCtaContcod " _
'        & "FROM   BalanceTemp as z JOIN " & sCentralCom & "CtaCont c ON z.cCtaContCod LIKE c.cCtaContCod + '%' " _
'        & "WHERE  cBalanceCate = '" & pnCateBala & "' and cBalanceTipo = '" & pnMoneda & "' " _
'        & "GROUP BY c.cCtaContCod, c.cCtaContDesc, cBalanceTipo " _
'        & "ORDER BY c.cCtaContCod "
  
   psSql = "exec stp_ins_BCient '" & Format(pdFecha, "yyyymmdd") & "','" & psCodigo & "','" & pnCateBala & "','" & pnMoneda & "' " 'PEAC 20210420
  
  dbConec.CommadTimeOut = 7200
  dbConec.Ejecutar psSql

If pnMoneda = "6" Then
   '69
'   psSql = "UPDATE BCIENT Set nSaldoIniImporte = nSaldoIniImporte * -1, nSaldoFinImporte = nSaldoFinImporte * -1 WHERE cCtaContCodAnt = '69' and cMoneda = '" & pnMoneda & "' and dFecha = '" & Format(pdFecha, gsFormatoFecha) & "' "
'   dbConec.Ejecutar psSql
'
'   psSql = "DELETE BCIENT Where cMoneda =  '0' and cCtaContCodAnt like '69' and dFecha = '" & Format(pdFecha, gsFormatoFecha) & "' "
'   dbConec.Ejecutar psSql
'
'   psSql = "Insert BCient SELECT dFecha, LEFT(cCtaContCod,9)+'69000000000000000000', Sum(nSaldoIniImporte ), Sum(nDebe), Sum(nHaber), Sum(nSaldoFinImporte), '0', '69' FROM BCient Where cCtaContCodAnt like '69' and dFecha = '" & Format(pdFecha, gsFormatoFecha) & "' GROUP BY dFecha, LEFT(cCtaContCod,9)+'69000000000000000000' "
'   dbConec.Ejecutar psSql

   '39
'   psSql = "UPDATE BCIENT Set nSaldoIniImporte = nSaldoIniImporte * -1, nSaldoFinImporte = nSaldoFinImporte * -1 WHERE cCtaContCodAnt = '39' and cMoneda = '" & pnMoneda & "' and dFecha = '" & Format(pdFecha, gsFormatoFecha) & "' "
'   dbConec.Ejecutar psSql

'   psSql = "DELETE BCIENT Where cMoneda =  '0' and cCtaContCodAnt like '39' and dFecha = '" & Format(pdFecha, gsFormatoFecha) & "' "
'   dbConec.Ejecutar psSql
'
'   psSql = "Insert BCient SELECT dFecha, LEFT(cCtaContCod,9)+'39000000000000000000', Sum(nSaldoIniImporte ), Sum(nDebe), Sum(nHaber), Sum(nSaldoFinImporte), '0', '39' FROM BCient Where cCtaContCodAnt like '39' and dFecha = '" & Format(pdFecha, gsFormatoFecha) & "' GROUP BY dFecha, LEFT(cCtaContCod,9)+'39000000000000000000' "
'   dbConec.Ejecutar psSql

End If


Exit Sub
InsertaBCientErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBCient Method")
End Sub

Public Sub ActualizaBCient(pdFecha As Date, psCodigo, pnCateBala As Integer, pnMoneda As Integer)
On Error GoTo ActualizaBCientErr
   
'   psSql = "UPDATE BCient Set cCtaContCod = '" & psCodigo & "' + cCtaContCodAnt + '0" & pnMoneda & "00000000000000000' " _
'       & " WHERE datediff(dd,dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and " _
'       & "       LEN(ltrim(rtrim(cCtacontcodant))) = 1 and cMoneda = '" & pnMoneda & "'"
'   dbConec.Ejecutar psSql
'
'   psSql = "UPDATE BCient Set cCtaContCod = '" & psCodigo & "' + cCtaContCodAnt + '" & pnMoneda & "00000000000000000' " _
'       & " WHERE datediff(dd,dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and " _
'       & " LEN(ltrim(rtrim(cCtacontcodant))) = 2 and cMoneda = '" & pnMoneda & "'"
'   dbConec.Ejecutar psSql
   
psSql = "exec stp_upd_BCient '" & Format(pdFecha, "yyyymmdd") & "','" & psCodigo & "','" & pnMoneda & "' "
dbConec.Ejecutar psSql
   
Exit Sub
ActualizaBCientErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ActualizaBCient Method")
End Sub


Public Sub EliminaBalanceSectorial(pdFecha As Date, Optional psCodigo As String = "", Optional lbEjecutarBatch As Boolean = False)
On Error GoTo EliminaBalanceErr
   psSql = "DELETE BalanceSectorial " _
         & "WHERE nBalanceMes = " & Month(pdFecha) & " and nBalanceAnio = " & Year(pdFecha) & IIf(psCodigo = "", "", " and cCodigo = '" & psCodigo & "'")
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

'ALPA 20100320*********************************************************
Public Function ObtenerDatosCtaCtesAhorroAyB(psFecha As String, psOpeTpo As String, psAgeCod As String) As Currency
On Error GoTo EliminaBalanceErr
   Dim oConec As DConecta
   Dim rs As ADODB.Recordset
   Dim nSumaPorAgencia As Currency
   Set oConec = New DConecta
   Set rs = New ADODB.Recordset
   oConec.AbreConexion
   psSql = "exec stp_sel_ObtenerDatosCtaCtesAhorroAyB  '" & psFecha & "','" & psOpeTpo & "','" & psAgeCod & "'"
   Set rs = oConec.CargaRecordSet(psSql)
   oConec.CierraConexion
   If rs.BOF Or rs.EOF Then
        ObtenerDatosCtaCtesAhorroAyB = 0
        Exit Function
   End If
   Do While Not rs.EOF
        nSumaPorAgencia = IIf(IsNull(rs!nSaldo), 0, rs!nSaldo)
        rs.MoveNext
   Loop
   oConec.CierraConexion
   ObtenerDatosCtaCtesAhorroAyB = nSumaPorAgencia
Exit Function
EliminaBalanceErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:EliminaBalance Method")
End Function
'*************************************************************************
Public Function CargaRABE() As Recordset
    Dim psSql As String
    psSql = " exec B2_stp_prueba_RABE"
   Set CargaRABE = dbConec.CargaRecordSet(psSql)
End Function

Public Sub InsertaBalanceTmpSaldosFecha(pnTipoBala As Integer, pnMoneda As Integer, psFecha As String, Optional lbConsolCta As Boolean = False, Optional pdFechador As Date)
Dim sCampoCta As String
On Error GoTo InsertaBalanceTmpSaldosFechaErr
   If pnMoneda = 0 And lbConsolCta Then
      sCampoCta = "substring(Cta.cCtaContCod,1,2) + '0' + substring(Cta.cCtaContCod,4,20)"
      psSql = " INSERT Into BalanceTemp " _
         & " SELECT '" & pnTipoBala & "','" & pnMoneda & "', " & sCampoCta & ", SUM(Cta.nCtaSaldoImporte), 0,0,cls.cCtaCaracter,'1'  " _
         & " FROM   CtaSaldo Cta JOIN " & sCentralCom & "CtaContClase cls ON cta.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' " _
         & " WHERE  cta.nCtaSaldoImporte <> 0 and Cta.dCtaSaldofecha in (SELECT MAX(Cta2.dCtaSaldofecha) From Ctasaldo Cta2 where cta2.cCtaContCod = Cta.cCtacontCod and Cta2.dCtaSaldofecha <= '" & psFecha & "') " & IIf(pnMoneda = 0, "", "and ( substring(cCtaContCod,3,1) = '" & pnMoneda & "'" & ")") _
         & " and DateDiff(d,dCtaSaldoFecha," & Format(pdFechador, "YYYY/MM/DD") & ")>=0" _
         & " GROUP BY " & sCampoCta & ", cls.cCtaCaracter"
   Else
      sCampoCta = "Cta.cCtaContCod"
      psSql = " INSERT Into BalanceTemp " _
         & " SELECT '" & pnTipoBala & "','" & pnMoneda & "', " & sCampoCta & ", Cta.nCtaSaldoImporte, 0,0,cls.cCtaCaracter,'1'  " _
         & " FROM   CtaSaldo Cta JOIN " & sCentralCom & "CtaContClase cls ON cta.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' " _
         & " WHERE  cta.nCtaSaldoImporte <> 0 and Cta.dCtaSaldofecha in (SELECT MAX(Cta2.dCtaSaldofecha) From Ctasaldo Cta2 where cta2.cCtaContCod = Cta.cCtacontCod and Cta2.dCtaSaldofecha <= '" & psFecha & "') " & IIf(pnMoneda = 0, "", "and ( substring(Cta.cCtaContCod,3,1) = '" & pnMoneda & "'" & ")") _
         & " and DateDiff(d,dCtaSaldoFecha," & Format(pdFechador, "YYYY/MM/DD") & ")>=0"
   End If
         
   dbConec.Ejecutar psSql
Exit Sub
InsertaBalanceTmpSaldosFechaErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceTmpSaldos Method")
End Sub

Public Sub InsertaCtaContSaldoDiario(psCtaContCod As String, pdFecDiaProceso As Date, psOpeCod As String, pnSaldoDiario As Currency)
On Error GoTo InsertaCtaContSaldoDiarioErr
   psSql = "exec stp_ins_CtaContSaldoDiario '" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psOpeCod & "'," & pnSaldoDiario
   dbConec.Ejecutar psSql

Exit Sub
InsertaCtaContSaldoDiarioErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Sub

Public Function ObtenerCtaContSaldoBalanceDiario(psCtaContCod As String, pdFecDiaProceso As Date, psMoneda As String, pnTipoCambio As Currency) As Currency
On Error GoTo ObtenerCtaContSaldoBalanceDiarioErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ObtenerDatosBalanceDiario '" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'," & pnTipoCambio
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = oRs!nSaldoFinImporte
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerCtaContSaldoBalanceDiario = nSaldo
Exit Function
ObtenerCtaContSaldoBalanceDiarioErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function


Public Function ListarSaldoCtas(psMoneda As String, psTpoCtas As String, psCuadro As String, pdFecDiaProceso As Date) As ADODB.Recordset
On Error GoTo ListarSaldoCtasErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ListarInformeEstadisticoBancosCajas '" & psMoneda & "','" & psTpoCtas & "','" & psCuadro & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   oConec.CierraConexion
   Set ListarSaldoCtas = oRs
Exit Function
ListarSaldoCtasErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Sub InsertaSaldoCtas(psAgeCod As String, psCodigoBancoAgeCod As String, psMoneda As String, psTpoCta As String, psNroCuadro As String, pdFecDiaProceso As Date, pnSaldoDiario As Currency)
On Error GoTo InsertaSaldoCtasErr
   psSql = "exec stp_ins_InformeEstadisticoBancosCajasDet '" & psAgeCod & "','" & psCodigoBancoAgeCod & "','" & psMoneda & "','" & psTpoCta & "','" & psNroCuadro & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'," & pnSaldoDiario
   dbConec.Ejecutar psSql
Exit Sub
InsertaSaldoCtasErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Sub

Public Function ObtenerCtaContSaldoDiario(psCtaContCod As String, pdFecDiaProceso As Date) As Currency
On Error GoTo ObtenerCtaContSaldoDiarioErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_CtaContSaldoDiario '" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = oRs!nSaldo
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerCtaContSaldoDiario = nSaldo
Exit Function
ObtenerCtaContSaldoDiarioErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Function ObtenerSumaInformeEstadisticoBancosCajas(pdFecDiaProceso As Date, psMoneda As String, psTpoCta As String) As Currency
On Error GoTo ObtenerSumaInformeEstadisticoBancosCajasErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SumInformeEstadisticoBancosCajasDET '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "','" & psTpoCta & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = oRs!nSaldo
           oRs.MoveNext
      Loop
   Else
        nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSumaInformeEstadisticoBancosCajas = nSaldo
Exit Function
ObtenerSumaInformeEstadisticoBancosCajasErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Function ObtenerCtaContBalanceMensual(psCtaContCod As String, pdFecDiaProceso As Date, psMoneda As String, pnTipoCambio As Currency) As Currency
On Error GoTo ObtenerCtaContBalanceMensualErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ObtenerDatosBalanceMensual '" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'," & pnTipoCambio
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = oRs!nSaldoFinImporte
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerCtaContBalanceMensual = nSaldo
Exit Function
ObtenerCtaContBalanceMensualErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function
'EJVG20130430 ***
Public Function ObtenerCtaContBalanceMensual2(ByVal psCtaContCod As String, ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, ByVal psBalanceCate As String, ByVal pnTipoCambio As Currency, ByVal pbMostrarSoles As Boolean) As Currency
    On Error GoTo ErrObtenerCtaContBalanceMensual2
    Dim oRs As New ADODB.Recordset
    Dim oConec As New DConecta
    oConec.AbreConexion
    psSql = "exec stp_sel_ObtenerDatosBalanceMensual2 '" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "','" & psBalanceCate & "'," & pnTipoCambio & "," & IIf(pbMostrarSoles, 1, 0)
    Set oRs = oConec.CargaRecordSet(psSql)
    If Not oRs.EOF And Not oRs.EOF Then
        ObtenerCtaContBalanceMensual2 = oRs!nSaldoFinImporte
    End If
    oConec.CierraConexion
    Set oConec = Nothing
    Set oRs = Nothing
    Exit Function
ErrObtenerCtaContBalanceMensual2:
    Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerCtaContBalanceMensual2 Method")
End Function
'END EJVG *******
Public Function ObtenerSaldoAdeudadoMyP(ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, ByVal psPlazo As String, ByVal psNacional As String) As Currency
On Error GoTo ObtenerSaldoAdeudadoMyPErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_MontoAdeudadosxPlazoyMoneda '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "','" & psPlazo & "','" & psNacional & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = oRs!CortoCap '+ oRs!CortoInt
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSaldoAdeudadoMyP = nSaldo
Exit Function
ObtenerSaldoAdeudadoMyPErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Function ObtenerSaldoBancosPF(ByVal pdFecDiaProceso As Date, ByVal psPersoneria As String, ByVal psMoneda As String) As Currency
On Error GoTo ObtenerSaldoBancosPFErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_obtenerSaldosPlazoFijoBancos '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psPersoneria & "','" & psMoneda & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = oRs!nSdoCnt
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSaldoBancosPF = nSaldo
Exit Function
ObtenerSaldoBancosPFErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function
Public Function ObtenerSaldoCtasRestringidas(ByVal pdFecDiaProceso As Date, ByVal psMoneda As String) As Currency
On Error GoTo ObtenerSaldoCtasRestringidasErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_obtenerCuentasRestringidasAnex15A '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = oRs!nSaldCnt
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSaldoCtasRestringidas = nSaldo
Exit Function
ObtenerSaldoCtasRestringidasErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Function ObtenerSaldoBancosAhorro(ByVal pdFecDiaProceso As Date, ByVal psPersoneria As String, ByVal psMoneda As String, ByVal pnLogCaja As Integer) As Currency
On Error GoTo ObtenerSaldoBancosAhorroErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_obtenerSaldosAhorrosBancos '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psPersoneria & "','" & psMoneda & "'," & pnLogCaja
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = oRs!nSdoCnt
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSaldoBancosAhorro = nSaldo
Exit Function
ObtenerSaldoBancosAhorroErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function
Public Function ObtenerFindeMes(ByVal pbFecha As Date) As Boolean
    Dim bLogicoFecha As Boolean
    Dim lnDiaFebrero As Integer
    If Month(pbFecha) = "01" Or Month(pbFecha) = "03" Or Month(pbFecha) = "05" Or Month(pbFecha) = "07" Or Month(pbFecha) = "08" Or Month(pbFecha) = "10" Or Month(pbFecha) = "12" Then
        If Day(pbFecha) = "31" Then
            bLogicoFecha = True
        Else
            bLogicoFecha = False
        End If
    ElseIf Month(pbFecha) = "02" Then
        If (Year(pbFecha) Mod 4) = 0 Then
            lnDiaFebrero = 28
        Else
            lnDiaFebrero = 29
        End If
        If Day(pbFecha) = lnDiaFebrero Then
            bLogicoFecha = True
        Else
            bLogicoFecha = False
        End If
    Else
        If Day(pbFecha) = "30" Then
            bLogicoFecha = True
        Else
            bLogicoFecha = False
        End If
    End If
    ObtenerFindeMes = bLogicoFecha
End Function

Public Function ObtenerSaldoInformeEstadisticoBancosBancos(ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, ByVal psBancos As String) As Currency
On Error GoTo ObtenerSaldoInformeEstadisticoBancosBancosErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SumInformeEstadisticoBancosCajasDETxBanco '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "','" & psBancos & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSaldoInformeEstadisticoBancosBancos = nSaldo
Exit Function
ObtenerSaldoInformeEstadisticoBancosBancosErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Function ObtenerDiasdeMes(ByVal pbFecha As Date) As Integer
    Dim nDias As Integer
    Dim lnDiaFebrero As Integer
    If Month(pbFecha) = "01" Or Month(pbFecha) = "03" Or Month(pbFecha) = "05" Or Month(pbFecha) = "07" Or Month(pbFecha) = "08" Or Month(pbFecha) = "10" Or Month(pbFecha) = "12" Then
            nDias = 31
    ElseIf Month(pbFecha) = "02" Then
        If (Year(pbFecha) Mod 4) = 0 Then
            nDias = 28
        Else
            nDias = 29
        End If
    Else
        nDias = 30
    End If
    ObtenerDiasdeMes = nDias
End Function

Public Function ObtenerSumaCtaContSaldoDiario(psCtaContCod As String, pdFecDiaProceso As Date) As Currency
On Error GoTo ObtenerSumaCtaContSaldoDiarioErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SumCtaContSaldoDiario '" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSumaCtaContSaldoDiario = nSaldo
Exit Function
ObtenerSumaCtaContSaldoDiarioErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Function ObtenerSumaValorRepresentativos(ByVal pdFecha As Date, ByVal TipoInstrumento As String) As Currency
Dim lsSql As String
Dim oCon As New DConecta
Dim oRs As New ADODB.Recordset
Dim lnSuma As Currency

On Error GoTo ErrobtenerSumaValorRepresentativos
oCon.AbreConexion
lnSuma = 0
lsSql = "exec sp_ERS0792016_sel_ListaTipoInstrumento '" & Format(pdFecha, "yyyymmdd") & "','" & TipoInstrumento & "'"
oCon.AbreConexion
Set oRs = oCon.CargaRecordSet(lsSql)

    If Not (oRs.BOF And oRs.EOF) Then
        Do While Not oRs.EOF
        lnSuma = lnSuma + IIf(IsNull(oRs!nValorRazonable), 0, oRs!nValorRazonable)
        oRs.MoveNext
        Loop
        Else
        lnSuma = 0
    End If
oCon.CierraConexion
    ObtenerSumaValorRepresentativos = lnSuma
Exit Function
ErrobtenerSumaValorRepresentativos:
    Call RaiseError(MyUnhandledError, "Recupera Valores Representativos")
End Function '****NAGL ERS079-2016 20170407
'****ANPS ObtenerSumaValorBCRP
Public Function ObtenerSumaValorBCRP(ByVal pdFecha As Date, ByVal Moneda As String, ByVal Cod As Integer) As Currency
Dim lsSql As String
Dim oCon As New DConecta
Dim oRs As New ADODB.Recordset
Dim lnSuma As Currency

On Error GoTo ObtenerSumaValorBCRP
oCon.AbreConexion
lnSuma = 0
lsSql = "exec sp_ObtenervalorInverciones '" & Format(pdFecha, "yyyymmdd") & "','" & Moneda & "'," & Cod & ""
oCon.AbreConexion
Set oRs = oCon.CargaRecordSet(lsSql)

    If Not (oRs.BOF And oRs.EOF) Then
      
        lnSuma = IIf(IsNull(oRs!nImporte), 0, oRs!nImporte)
        
    Else
        lnSuma = 0
    End If
oCon.CierraConexion
    ObtenerSumaValorBCRP = lnSuma
Exit Function
ObtenerSumaValorBCRP:
    Call RaiseError(MyUnhandledError, "Recupera Valores Representativos BCRP")
End Function '****ANPS
Public Function ObtenerPosicionAcumMesAnterior(ByVal pdFecha As Date, ByVal psMoneda As String) As Currency
Dim lsSql As String
Dim oCon As New DConecta
Dim oRs As New ADODB.Recordset
Dim lnSaldo As Currency

On Error GoTo ErrObtenerPosicionAcumMesAnterior
oCon.AbreConexion
lnSaldo = 0
lsSql = "exec stp_sel_PosicionAcumuladoMesAnterior '" & Format(pdFecha, "yyyymmdd") & "','" & psMoneda & "'"
oCon.AbreConexion
Set oRs = oCon.CargaRecordSet(lsSql)
    If Not (oRs.BOF And oRs.EOF) Then
        lnSaldo = oRs!nSaldo
        oRs.MoveNext
        Else
        lnSaldo = 0
    End If
oCon.CierraConexion
    ObtenerPosicionAcumMesAnterior = lnSaldo
Exit Function
ErrObtenerPosicionAcumMesAnterior:
    Call RaiseError(MyUnhandledError, "Recupera Posicion Acumulada Mes Anterior")
End Function '****NAGL 20170422

Public Function ObtenerSumaOperacionesReportesVcto30(ByVal pdFecha As Date, ByVal psMoneda As String) As Currency
Dim lsSql As String
Dim oCon As New DConecta
Dim oRs As New ADODB.Recordset
Dim lnSuma As Currency

On Error GoTo ErrobtenerSumaOperacionesReportesVcto30
oCon.AbreConexion
lnSuma = 0
lsSql = "exec sp_ERS0792016_sel_OperacionesReportesvct30dias '" & Format(pdFecha, "yyyymmdd") & "','" & psMoneda & "'"
oCon.AbreConexion
Set oRs = oCon.CargaRecordSet(lsSql)
   
   If Not (oRs.BOF And oRs.EOF) Then
   Do While Not oRs.EOF
   lnSuma = lnSuma + IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
   oRs.MoveNext
   Loop
   Else
   lnSuma = 0
   End If
oCon.CierraConexion
   ObtenerSumaOperacionesReportesVcto30 = lnSuma
Exit Function
ErrobtenerSumaOperacionesReportesVcto30:
    Call RaiseError(MyUnhandledError, "Operaciones de Reportes con vcto 30 dias")
End Function '****NAGL ERS079-2016 20170407

Public Function ObtenerActivosLiquidosReporte15A(ByVal pdFecha As Date, ByVal psMoneda As String, ByVal cReporte As String, ByVal cCodigoText As String) As Currency
Dim lsSql As String
Dim oCon As New DConecta
Dim oRs As ADODB.Recordset
Set oRs = New ADODB.Recordset
Dim lnSaldo As Currency
On Error GoTo ErroObtenerActivoLiquidosReporte15A
oCon.AbreConexion
lnSaldo = 0
lsSql = "exec stp_ERS0792016_sel_Reporte15DetalleDiario '" & Format(pdFecha, "yyyymmdd") & "','" & psMoneda & "', '" & cReporte & "','" & cCodigoText & "'"
oCon.AbreConexion
Set oRs = oCon.CargaRecordSet(lsSql)
If Not (oRs.BOF And oRs.EOF) Then
    Do While Not oRs.EOF
    lnSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
    oRs.MoveNext
Loop
End If
oCon.CierraConexion
   ObtenerActivosLiquidosReporte15A = lnSaldo
Exit Function
ErroObtenerActivoLiquidosReporte15A:
    Call RaiseError(MyUnhandledError, "Activos Liquidos Reporte 15A")
End Function '****NAGL ERS079-2016 20170407
'JIPR20201023
Public Function ObtenerValorAnexo(ByVal pParProd As Integer, ByVal pParCod As Integer) As Currency
Dim lsSql As String
Dim oCon As New DConecta
Dim oRs As ADODB.Recordset
Set oRs = New ADODB.Recordset
Dim nParValor As Currency
On Error GoTo ErroObtenerValorAnexo
oCon.AbreConexion
nParValor = 0
lsSql = "exec stp_sel_ObtenerValorAnexo " & pParProd & "," & pParCod & ""
oCon.AbreConexion
Set oRs = oCon.CargaRecordSet(lsSql)
If Not (oRs.BOF And oRs.EOF) Then
    Do While Not oRs.EOF
    nParValor = IIf(IsNull(oRs!nParValor), 0, oRs!nParValor)
    oRs.MoveNext
Loop
End If
oCon.CierraConexion
   ObtenerValorAnexo = nParValor
Exit Function
ErroObtenerValorAnexo:
    Call RaiseError(MyUnhandledError, " Valores Anexos")
End Function 'JIPR20201023

Public Function ObtenerMontosyTasasSubastaTesoroPublico(ByVal pdFecha As Date) As Recordset
On Error GoTo ErroObtenerMontosyTasasSubastaTesoroPublico
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec sp_ERS0792016_sel_MontosyTasasSubastaTesoroPublica '" & Format(pdFecha, "YYYY/MM/DD") & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerMontosyTasasSubastaTesoroPublico = oRs
   oConec.CierraConexion
Exit Function
ErroObtenerMontosyTasasSubastaTesoroPublico:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerMontosyTasasSubastaTesoroPublico Method")
End Function '****NAGL ERS079-2016 20170407

Public Function ObtenerPermiteGenerarfromBitacora15A_15B(ByVal pdFecha As Date, ByVal pdFechaSist As String) As String
Dim lsSql As String
Dim oCon As New DConecta
Dim oRs As ADODB.Recordset
Set oRs = New ADODB.Recordset
Dim lsPermite As String
On Error GoTo ErroObtenerPermiteGenerarfromBitacora15A_15B
oCon.AbreConexion
lsSql = "exec stp_sel_GetPermiteGenerarAnx15A_15BfromBitacora '" & Format(pdFecha, "yyyymmdd") & "','" & Format(pdFechaSist, "yyyymmdd") & "'"
oCon.AbreConexion
Set oRs = oCon.CargaRecordSet(lsSql)
If Not (oRs.BOF And oRs.EOF) Then
    Do While Not oRs.EOF
        lsPermite = oRs!PermiteUsoBitacora
        oRs.MoveNext
    Loop
End If
oCon.CierraConexion
   ObtenerPermiteGenerarfromBitacora15A_15B = lsPermite
Exit Function
ErroObtenerPermiteGenerarfromBitacora15A_15B:
    Call RaiseError(MyUnhandledError, "Generación desde Bitácora para Anx15A_15B")
End Function '****NAGL ERS002-2017 20170815

Public Sub InsertaDetallaReporte15A(ByVal pnNroItem As Integer, ByVal pdFecDiaProceso As Date, psMoneda As String, ByVal pnSaldoDiario As Currency, Optional ByVal pnVisibleC As Integer = 1, Optional psCodigoText As String = "000", Optional psReporte As String = "A1")
On Error GoTo InsertaDetallaReporte15AErr
   psSql = "exec stp_ins_Reporte15DetalleDiario " & pnNroItem & ",'" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'," & pnSaldoDiario & "," & pnVisibleC & ",'" & psCodigoText & "', '" & psReporte & "'"
   dbConec.Ejecutar psSql
Exit Sub
InsertaDetallaReporte15AErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaDetallaReporte15A Method")
End Sub
Public Function ObtenerListaReporte15B(ByVal pdFecDiaProceso As Date, psMoneda As String, Optional pnContador As Integer = 1, Optional cReporte As String = "A1", Optional ByRef psFechaInicial As Date = "1900/01/01") As ADODB.Recordset
On Error GoTo ObtenerListaReporte15BErr

   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_Reporte15DetalleDiario '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "','" & cReporte & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If pnContador < 32 Then
    If (oRs.BOF Or oRs.EOF) Then
         pnContador = pnContador + 1
         pdFecDiaProceso = DateAdd("d", -1, pdFecDiaProceso)
         Set oRs = ObtenerListaReporte15B(pdFecDiaProceso, psMoneda, pnContador, cReporte, psFechaInicial)
    End If
   End If
    'Call InsertaDetallaReporte15ADe15C(pdFecDiaProceso, psMoneda, "A1", psFechaInicial)'Comentado by NAGL 20190621
   Set ObtenerListaReporte15B = oRs
   oConec.CierraConexion
Exit Function
ObtenerListaReporte15BErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function
'ALPA20131210***********************************************************************************************************************************************************
Public Sub InsertaDetallaReporte15ADe15C(ByVal pdFecDiaProceso As Date, psMoneda As String, psReporte As String, Optional psFechaInicial As Date = "1900/01/01")
On Error GoTo InsertaDetallaReporte15ADe15CErr
   psSql = "exec stp_ins_Reporte15ADetalleDiarioDe15C '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "','" & psReporte & "','" & Format(psFechaInicial, "YYYY/MM/DD") & "'"
   'dbConec.Ejecutar psSql
Exit Sub
InsertaDetallaReporte15ADe15CErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaDetallaReporte15ADe15C Method")
End Sub
'***********************************************************************************************************************************************************************
'JACA 20111017****************************************************************
Public Function ObtenerDatosAnexo12_II(ByVal psAnio As String, ByVal psMes As String) As Recordset
On Error GoTo ErrorRep
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   
   oConec.AbreConexion
   psSql = "exec stp_sel_Anexo12_II '" & psAnio & "','" & psMes & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerDatosAnexo12_II = oRs
   oConec.CierraConexion
Exit Function
ErrorRep:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerDatosAnexo12_II Method")

End Function

'JACA END********************************************************************

'***Agregado por ELRO el 20111222, según Acta N° 352-2011/TI-D

Public Function recuperarPatrimonioEfectivoEleccion(ByVal pnTipo As Integer, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal pdFecha As Date, ByVal pdFechaSist As Date) As ADODB.Recordset
On Error GoTo recuperarPatrimonioEfectivoEleccionError
   Dim oDConecta As DConecta
   Set oDConecta = New DConecta
   oDConecta.AbreConexion
   psSql = "exec stp_sel_RecuperarPatrimonioEfectivoEleccion " & pnTipo & ", " & pnMes & ", " & pnAnio & ", '" & Format(pdFecha, "yyyymmdd") & "', '" & Format(pdFechaSist, "yyyymmdd") & "' "
   Set recuperarPatrimonioEfectivoEleccion = oDConecta.CargaRecordSet(psSql)
   oDConecta.CierraConexion
Exit Function
recuperarPatrimonioEfectivoEleccionError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:recuperarPatrimonioEfectivo Eleccion Method")
End Function 'NAGL ERS 002-2017 20170802

Public Function ObtenerTipoCambioCierreNew(ByVal pdFecha As Date, Optional psTipo As String = "") As Currency
Dim lsSql As String
Dim oCon As New DConecta
Dim oRs As New ADODB.Recordset
Dim lnMonto As Currency

On Error GoTo ErrObtenerTipoCambioCierreNew
oCon.AbreConexion
    lnMonto = 0
        lsSql = "exec stp_sel_ObtieneTipoCambioNew '" & Format(pdFecha, "yyyymmdd") & "','" & psTipo & "'" 'Interviene tbn el psTipo = "TipoAct"
    oCon.AbreConexion
Set oRs = oCon.CargaRecordSet(lsSql)
   
   If Not (oRs.BOF And oRs.EOF) Then
        Do While Not oRs.EOF
            lnMonto = lnMonto + IIf(IsNull(oRs!TipoCambio), 0, oRs!TipoCambio)
            oRs.MoveNext
        Loop
   Else
        lnMonto = 0
   End If
oCon.CierraConexion
   ObtenerTipoCambioCierreNew = lnMonto
Exit Function
ErrObtenerTipoCambioCierreNew:
    Call RaiseError(MyUnhandledError, "Obtiene Tipo De Cambio")
End Function 'NAGL ERS 002-2017 20170914

Public Function recuperarPatrimonioEfectivo() As ADODB.Recordset
On Error GoTo recuperarPatrimonioEfectivoError
   Dim oDConecta As DConecta
   Set oDConecta = New DConecta
   
   oDConecta.AbreConexion
   psSql = "exec stp_sel_RecuperarPatrimonioEfectivo"
   Set recuperarPatrimonioEfectivo = oDConecta.CargaRecordSet(psSql)
   oDConecta.CierraConexion
Exit Function
recuperarPatrimonioEfectivoError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:recuperarPatrimonioEfectivo Method")
End Function

Public Function registrarPatrimonioEfectivo(ByVal pnSaldo As Currency, ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal pcMov As String, Optional psTipo As String = "") As Boolean
    
On Error GoTo registrarPatrimonioEfectivoError
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta
    
    registrarPatrimonioEfectivo = False
    oDConecta.AbreConexion
    
    psSql = "stp_ins_RegistrarPatrimonioEfectivo " & pnSaldo & ", " & pnMes & ", " & pnAnio & ", '" & pcMov & "', '" & psTipo & "'" 'NAGL Agregó psTipo 20190118
    oDConecta.Ejecutar (psSql)
    registrarPatrimonioEfectivo = True
    oDConecta.CierraConexion
Exit Function
registrarPatrimonioEfectivoError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:registrarPatrimonioEfectivo Method")
End Function

Public Function actualizarPatrimonioEfectivo(ByVal pIdPatrimonio As Integer, _
                                             ByVal pnSaldo As Currency, _
                                             ByVal pnMes As Integer, _
                                             ByVal pnAnio As Integer, _
                                             ByVal pcMov As String) As Boolean
On Error GoTo actualizarPatrimonioEfectivoError
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta
    
    actualizarPatrimonioEfectivo = False
    oDConecta.AbreConexion
    
    psSql = "stp_upd_ActualizarPatrimonioEfectivo " & pIdPatrimonio & ", " & pnSaldo & ", " & pnMes & ", " & pnAnio & ", '" & pcMov & "'"
    oDConecta.Ejecutar (psSql)
    actualizarPatrimonioEfectivo = True
    oDConecta.CierraConexion
Exit Function
actualizarPatrimonioEfectivoError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:actualizarPatrimonioEfectivo Method")
End Function

Public Function eliminarPatrimonioEfectivo(ByVal pIdPatrimonio As Integer) As Boolean
On Error GoTo eliminarPatrimonioEfectivoError
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta
    
    eliminarPatrimonioEfectivo = False
    oDConecta.AbreConexion
    
    psSql = "stp_del_EliminarPatrimonioEfectivo " & pIdPatrimonio & ""
    oDConecta.Ejecutar (psSql)
    eliminarPatrimonioEfectivo = True
    oDConecta.CierraConexion
Exit Function
eliminarPatrimonioEfectivoError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:eliminarPatrimonioEfectivo Method")
   
End Function

Public Sub eliminarBalanceTempHist(ByVal psFechaCorte As String)
Dim sSQL As String

On Error GoTo eliminarBalanceTempHistErr
    sSQL = "stp_del_EliminarBalanceTempHist '" & Format(psFechaCorte, gsFormatoMovFecha) & "'"
    dbConec.Ejecutar sSQL
Exit Sub
eliminarBalanceTempHistErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:eliminarBalanceTempHist Method")
End Sub

Public Sub InsertaBalanceTempHistSaldos(ByVal psMov As String)
Dim sSQL As String
On Error GoTo InsertaBalanceTempHistSaldosErr
    sSQL = "exec stp_ins_RegistrarBalanceTempHist '" & psMov & "'"
   dbConec.Ejecutar sSQL
Exit Sub
InsertaBalanceTempHistSaldosErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceTempHistSaldos Method")
End Sub

'***Fin Agregado por ELRO*************************************
'ALPA 20121220************************************************
Public Function ObtenerSumaCarteraCreditos(pdFecDiaProceso As Date, psMoneda As String) As Currency
On Error GoTo ObtenerSumaCarteraCreditosErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_obtenerCarteraCredito '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = oRs!nMontoTotal
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSumaCarteraCreditos = nSaldo
Exit Function
ObtenerSumaCarteraCreditosErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Function ObtenerSumaFondeo(pdFecDiaProceso As Date, psMoneda As String) As Currency
On Error GoTo ObtenerSumaFondeoErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_fondeo '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = oRs!nSaldo
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSumaFondeo = nSaldo
Exit Function
ObtenerSumaFondeoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Function ObtenerSumaNuevaCarteraVencida(pdFecDiaProceso As Date, psMoneda As String) As Currency
On Error GoTo ObtenerSumaNuevaCarteraVencidaErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SaldoNuevaCarteraVencida '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = oRs!nSaldo
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSumaNuevaCarteraVencida = nSaldo
Exit Function
ObtenerSumaNuevaCarteraVencidaErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Function RecuperarCarteraxTipoCredito(ByVal pdFecha As Date) As ADODB.Recordset
On Error GoTo RecuperarCarteraxTipoCreditoError
   Dim oDConecta As DConecta
   Set oDConecta = New DConecta
   
   oDConecta.AbreConexion
   psSql = "exec stp_sel_obtenerCarteraCreditoxtipoCredito '" & Format(pdFecha, "YYYY/MM/DD") & "'"
   Set RecuperarCarteraxTipoCredito = oDConecta.CargaRecordSet(psSql)
   oDConecta.CierraConexion
Exit Function
RecuperarCarteraxTipoCreditoError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:recuperarPatrimonioEfectivo Method")
End Function


Public Function RecuperarCarteraxSectorEconomico(ByVal pdFecha As Date) As ADODB.Recordset
On Error GoTo RecuperarCarteraxSectorEconomicoError
   Dim oDConecta As DConecta
   Set oDConecta = New DConecta
   
   oDConecta.AbreConexion
   psSql = "exec stp_sel_obtenerCarteraCreditoxSectorEconomico '" & Format(pdFecha, "YYYY/MM/DD") & "'"
   Set RecuperarCarteraxSectorEconomico = oDConecta.CargaRecordSet(psSql)
   oDConecta.CierraConexion
Exit Function
RecuperarCarteraxSectorEconomicoError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:recuperarPatrimonioEfectivo Method")
End Function
'*****************************************************************
'EJVG20121112 ***
Public Function RecuperaInfoEstadColocBCRP(ByVal pnAnio As Integer, ByVal pnMes As Integer, ByVal psAgeCod As String) As ADODB.Recordset
    Dim oDConecta As New DConecta
    On Error GoTo ErrRecuperaInfoEstadColocBCRP
    oDConecta.AbreConexion
    psSql = "exec stp_sel_ReporteInfoEstadColocBCRP '" & Format(pnAnio, "0000") & "','" & Format(pnMes, "00") & "','" & psAgeCod & "'"
    Set RecuperaInfoEstadColocBCRP = oDConecta.CargaRecordSet(psSql)
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Exit Function
ErrRecuperaInfoEstadColocBCRP:
   Call RaiseError(MyUnhandledError, "DBalanceCont:recuperarPatrimonioEfectivo Method")
End Function
'END EJVG *******
'MIOL20130209 RFC138***
Public Function RecuperarPatrimonioEfectivoMesAnio(ByVal pnAnio As Integer, ByVal pnMes As Integer) As ADODB.Recordset
    Dim oDConecta As New DConecta
    On Error GoTo ErrrecuperarPatrimonioEfectivoMesAnio
    oDConecta.AbreConexion
    psSql = "exec stp_sel_PatrimonioMesAnio '" & Format(pnAnio, "0000") & "','" & Format(pnMes, "00") & "'"
    Set RecuperarPatrimonioEfectivoMesAnio = oDConecta.CargaRecordSet(psSql)
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Exit Function
ErrrecuperarPatrimonioEfectivoMesAnio:
   Call RaiseError(MyUnhandledError, "DBalanceCont:recuperarPatrimonioEfectivo Method")
End Function
'END MIOL *******
'ALPA 20130612********************************
Public Function RecuperarDepartamAgencia() As ADODB.Recordset
    Dim oDConecta As New DConecta
    On Error GoTo ErrRecuperarDepartamAgencia
    oDConecta.AbreConexion
    psSql = "exec stp_sel_DepartamAgencia "
    Set RecuperarDepartamAgencia = oDConecta.CargaRecordSet(psSql)
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Exit Function
ErrRecuperarDepartamAgencia:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ErrRecuperarDepartamAgencia Method")
End Function
Public Function ObtenerDispobiblesenSFN(psPersCod As String, psCtaContCod As String, pdFecDiaProceso As Date, nTipo As Integer) As Currency
On Error GoTo ObtenerDispobiblesenSFNErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_DispobiblesenSFN '" & psPersCod & "','" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'," & nTipo
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
           oRs.MoveNext
      Loop
   Else
        nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerDispobiblesenSFN = nSaldo
Exit Function
ObtenerDispobiblesenSFNErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Function ObtenerCtaSaldoDiario(psCtaContCod As String, pdFecDiaProceso As Date) As Currency
On Error GoTo ObtenerCtaSaldoDiarioErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_CtaSaldoDiario '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psCtaContCod & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
           oRs.MoveNext
      Loop
   Else
        nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerCtaSaldoDiario = nSaldo
Exit Function
ObtenerCtaSaldoDiarioErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Function ObtenerCantidadPersonal(psMes As String, psAnio As String) As Integer
On Error GoTo ObtenerCantidadPersonalErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nNroTrabajadores As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec STP_SEL_CANTTRABAJADORES '" & psMes & "','" & psAnio & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nNroTrabajadores = IIf(IsNull(oRs!nNroTrabajadores), 0, oRs!nNroTrabajadores)
           oRs.MoveNext
      Loop
   Else
        nNroTrabajadores = 0
   End If
   oConec.CierraConexion
   ObtenerCantidadPersonal = nNroTrabajadores
Exit Function
ObtenerCantidadPersonalErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerCantidadPersonal Method")
End Function

Public Function ObtenerPlazoFijoxRango(pdFecha As Date, pnFiltro As Integer, pnMoneda As Integer, pnRango As Long) As Double
On Error GoTo ObtenerPlazoFijoxRangoErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSumaTramos As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_PlazoFijoRangosPersoneria '" & Format(pdFecha, "YYYY/MM/DD") & "','" & pnFiltro & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   nSumaTramos = 0
   If Not oRs.BOF And Not oRs.EOF Then
      Do While Not oRs.EOF
            If pnRango >= oRs!Fin Then
                If pnMoneda = "1" Then
                  nSumaTramos = nSumaTramos + IIf(IsNull(oRs!Saldo_MN), 0, oRs!Saldo_MN)
                Else
                  nSumaTramos = nSumaTramos + IIf(IsNull(oRs!Saldo_ME), 0, oRs!Saldo_ME)
                End If
            End If
           oRs.MoveNext
      Loop
   Else
        nSumaTramos = 0
   End If
   oConec.CierraConexion
   ObtenerPlazoFijoxRango = nSumaTramos
Exit Function
ObtenerPlazoFijoxRangoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerPlazoFijoxRango:Method")
End Function
Public Function ObtenerCuentaHistorica(psCtaCodCod As String) As String
On Error GoTo ObtenerCuentaHistoricaErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim lsCtaContCodAnterior As String
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_RelacionCtaContHistxCta '" & psCtaCodCod & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        lsCtaContCodAnterior = IIf(IsNull(oRs!cCtaContCodHisto), "", oRs!cCtaContCodHisto)
           oRs.MoveNext
      Loop
   Else
        lsCtaContCodAnterior = psCtaCodCod
   End If
   oConec.CierraConexion
   ObtenerCuentaHistorica = lsCtaContCodAnterior
Exit Function
ObtenerCuentaHistoricaErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerCantidadPersonal Method")
End Function

Public Function ObtenerCtaContBalanceMensualxAgencia(ByVal psCtaContCod As String, ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, ByVal psBalanceCate As String, ByVal pnTipoCambio As Currency, ByVal pbMostrarSoles As Boolean, ByVal psAgeCod As String) As Currency
    On Error GoTo ErrObtenerCtaContBalanceMensualxAgencia
    Dim oRs As New ADODB.Recordset
    Dim oConec As New DConecta
    oConec.AbreConexion
    psSql = "exec stp_sel_ObtenerDatosBalanceMensual2AG '" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "','" & psBalanceCate & "'," & pnTipoCambio & "," & IIf(pbMostrarSoles, 1, 0) & ",'" & psAgeCod & "'"
    Set oRs = oConec.CargaRecordSet(psSql)
    If Not oRs.EOF And Not oRs.EOF Then
        ObtenerCtaContBalanceMensualxAgencia = IIf(IsNull(oRs!nSaldoFinImporte), 0, oRs!nSaldoFinImporte)
    End If
    oConec.CierraConexion
    Set oConec = Nothing
    Set oRs = Nothing
    Exit Function
ErrObtenerCtaContBalanceMensualxAgencia:
    Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerCtaContBalanceMensualxAgencia Method")
End Function

Public Function ObtenerSaldoAhorroxPersoneria(ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, ByVal psPersoneria As String) As Currency
On Error GoTo ObtenerSaldoAhorroxPersoneriaErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ReporteSaldosDiarios '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "','" & psPersoneria & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldCnt), 0, oRs!nSaldCnt)
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSaldoAhorroxPersoneria = nSaldo
Exit Function
ObtenerSaldoAhorroxPersoneriaErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerSaldoAhorroxPersoneria Method")
End Function

Public Function ObtenerParamEncDiarioxCodigo(ByVal psCodigo As String, Optional psMoneda As String = "", Optional psFecha As String = "") As Currency
On Error GoTo ObtenerParamEncDiarioxCodigoErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nValor As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ParamEncDiarioxCodigo '" & psCodigo & "','" & psMoneda & "','" & psFecha & "'" 'NAGL Agregó psMoneda, y pdFecha Según TIC1810110004
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nValor = IIf(IsNull(oRs!nValor), 0, oRs!nValor)
           oRs.MoveNext
      Loop
   Else
    nValor = 0
   End If
   oConec.CierraConexion
   ObtenerParamEncDiarioxCodigo = nValor
Exit Function
ObtenerParamEncDiarioxCodigoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerParamEncDiarioxCodigo Method")
End Function
'*******************************************************************************************
'ALPA20130822*******************************************************************************
Public Function ObtenerSaldoOverNight(ByVal psMoneda As String, ByVal pdFecDiaProceso As Date, ByVal pnTipo As Integer) As Currency
On Error GoTo ObtenerSaldoOverNightErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_OverNight '" & psMoneda & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & pnTipo & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
        oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSaldoOverNight = nSaldo
Exit Function
ObtenerSaldoOverNightErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerSaldoOverNight Method")
End Function
Public Function ObtenerSaldoDepositosGrandesAcreedores_Estado(ByVal psMoneda As String, ByVal pdFecDiaProceso As Date, ByVal pnTipoCambio As Currency) As Currency
On Error GoTo ObtenerSaldoDepositosGrandesAcreedores_EstadoErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SaldoDepositosGrandesAcreedores_Estado '" & psMoneda & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & pnTipoCambio & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
        oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSaldoDepositosGrandesAcreedores_Estado = nSaldo
Exit Function
ObtenerSaldoDepositosGrandesAcreedores_EstadoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerSaldoDepositosGrandesAcreedores_Estado Method")
End Function

Public Function ObtenerSaldoDepositosGrandesAcreedores_Otros(ByVal psMoneda As String, ByVal pdFecDiaProceso As Date, ByVal pnTipoCambio As Currency, Optional psTipoGA As String = "", Optional psTipo As String = "") As Currency
On Error GoTo ObtenerSaldoDepositosGrandesAcreedores_OtrosErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   
   psSql = "exec stp_sel_ERS006_2019_SaldoDepGrandesAcreedores_Otros'" & psMoneda & "','" & Format(pdFecDiaProceso, "yyyymmdd") & "','" & pnTipoCambio & "','" & psTipoGA & "','" & psTipo & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
      Do While Not oRs.EOF
        nSaldo = nSaldo + IIf(IsNull(oRs!Monto), 0, oRs!Monto)
        oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSaldoDepositosGrandesAcreedores_Otros = nSaldo
Exit Function
ObtenerSaldoDepositosGrandesAcreedores_OtrosErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerSaldoDepositosGrandesAcreedores_Otros Method")
End Function 'NAGL 20190627 Según Anx02_ERS006-2019

Public Function ListarSaldoDepGrandesAcreedores(ByVal psMoneda As String, ByVal pdFecDiaProceso As Date, ByVal pnTipoCambio As Currency, Optional psTipoGA As String = "", Optional psTipo As String = "") As ADODB.Recordset
    Dim oConec As New DConecta
    Dim psSql As String
    On Error GoTo ErroListarSaldoDepGrandesAcreedores
    
    psSql = "exec stp_sel_ERS006_2019_SaldoDepGrandesAcreedores_Otros'" & psMoneda & "','" & Format(pdFecDiaProceso, "yyyymmdd") & "','" & pnTipoCambio & "','" & psTipoGA & "','" & psTipo & "'"
    oConec.AbreConexion
    Set ListarSaldoDepGrandesAcreedores = oConec.CargaRecordSet(psSql)
    oConec.CierraConexion
    
    Exit Function
ErroListarSaldoDepGrandesAcreedores:
    Call RaiseError(MyUnhandledError, "Lista Depositos de Grandes Acreedores")
End Function '***********NAGL ERS006-2019 20190627

Public Function ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptaciones(ByVal pnNroCliente As String, ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, pnTipoCambio As Currency, Optional pnTipo As Integer = 0) As Currency
On Error GoTo ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptaciones " & pnNroCliente & ",'" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'," & pnTipoCambio & "," & pnTipo
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
        oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptaciones = nSaldo
Exit Function
ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerSaldoDepositosGrandesAcreedores_Estado Method")
End Function
Public Function recuperarPatrimonioEfectivoxFecha(ByVal pdFecha As Date) As Currency
On Error GoTo recuperarPatrimonioEfectivoxFechaError

   Dim oRs As ADODB.Recordset
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Dim oDConecta As DConecta
   Set oDConecta = New DConecta
   
   oDConecta.AbreConexion
   psSql = "exec stp_sel_RecuperarPatrimonioEfectivoxFecha '" & Format(pdFecha, "YYYY/MM/DD") & "'"
   Set oRs = oDConecta.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
        oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   recuperarPatrimonioEfectivoxFecha = nSaldo
   oDConecta.CierraConexion
Exit Function
recuperarPatrimonioEfectivoxFechaError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:recuperarPatrimonioEfectivoxFecha Method")
End Function
Public Function ObtenerCtaSaldoDiarioxMoneda(psCtaContCod As String, pdFecDiaProceso As Date, psMoneda As String) As Currency
On Error GoTo ObtenerCtaSaldoDiarioxMonedaErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_CtaSaldoDiarioxMoneda '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psCtaContCod & "','" & psMoneda & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
           oRs.MoveNext
      Loop
   Else
        nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerCtaSaldoDiarioxMoneda = nSaldo
Exit Function
ObtenerCtaSaldoDiarioxMonedaErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerCtaSaldoDiarioxMoneda Method")
End Function
Public Function obtenerSumaConsolidadoCtasxConcentracionFondos(ByVal pdFecha As Date, ByVal pnTpoMoneda As Integer, ByVal psListaTpoIfi As String, ByVal psListaCodPersIFiExcluidas As String) As Currency
    Dim lsSql As String
    Dim oCon As New DConecta
    Dim oRs As New ADODB.Recordset
    Dim lnSuma As Currency
On Error GoTo ErrobtenerSumaConsolidadoCtasxConcentracionFondos
    oCon.AbreConexion
    lnSuma = 0
    lsSql = "Exec stp_sel_ListaConsolCtaRptConcentracion '" & Format(pdFecha, "yyyymmdd") & "','" & pnTpoMoneda & "','" & psListaTpoIfi & "','" & psListaCodPersIFiExcluidas & "'"
    oCon.AbreConexion
    Set oRs = oCon.CargaRecordSet(lsSql)
    If Not (oRs.BOF Or oRs.EOF) Then
        Do While Not oRs.EOF
        lnSuma = lnSuma + IIf(IsNull(oRs!nSaldoCtaCorriente), 0, oRs!nSaldoCtaCorriente) + IIf(IsNull(oRs!nSaldoCtaAhorro), 0, oRs!nSaldoCtaAhorro) + IIf(IsNull(oRs!nSaldoDPFOVer), 0, oRs!nSaldoDPFOVer)
        oRs.MoveNext
        Loop
    End If
    oCon.CierraConexion
    obtenerSumaConsolidadoCtasxConcentracionFondos = lnSuma
    Exit Function
ErrobtenerSumaConsolidadoCtasxConcentracionFondos:
    Call RaiseError(MyUnhandledError, "Recupera Ctas Concentración")
End Function
Public Function ObtenerFondeoEncaje(ByVal pdFecha As Date, ByVal pnTipoCambio As Currency, Optional ByVal pnPlazo As Long = 99999) As ADODB.Recordset
    Dim lsSql As String
    Dim oCon As New DConecta
    Dim oRs As New ADODB.Recordset
    Dim lnSuma As Currency
On Error GoTo ErrObtenerFondeoEncaje
    oCon.AbreConexion
    lnSuma = 0
    lsSql = "Exec stp_sel_FondeoEncaje '" & Format(pdFecha, "yyyymmdd") & "'," & pnTipoCambio & "," & pnPlazo
    oCon.AbreConexion
    Set ObtenerFondeoEncaje = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    Exit Function
ErrObtenerFondeoEncaje:
    Call RaiseError(MyUnhandledError, "ObtenerFondeoEncaje")
End Function
Public Function ObtenerFondeoEncajeConsolidado(ByVal pdFecha As Date, ByVal pnTipoCambio As Currency) As ADODB.Recordset
    Dim lsSql As String
    Dim oCon As New DConecta
    Dim oRs As New ADODB.Recordset
On Error GoTo ErrObtenerFondeoEncajeConsolidado
    oCon.AbreConexion
    lsSql = "Exec stp_sel_FondeoEncajeConsolidado '" & Format(pdFecha, "yyyymmdd") & "'," & pnTipoCambio & ""
    Set ObtenerFondeoEncajeConsolidado = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    Exit Function
ErrObtenerFondeoEncajeConsolidado:
    Call RaiseError(MyUnhandledError, "ObtenerFondeoEncajeConsolidado")
End Function '***NAGL Según ERS006-2019 20190424

Public Function ObtenerObligInstRecMenora30Dias(ByVal pdFecha As Date, pdFechaBalanceDiario As Date, pnTipoCambBal As Currency, pnTipoCambioMes As Currency) As ADODB.Recordset
    Dim lsSql As String
    Dim oCon As New DConecta
    Dim oRs As New ADODB.Recordset
On Error GoTo ErrObtenerObligInstRecMenora30Dias
    oCon.AbreConexion
    lsSql = "Exec stp_sel_ObtieneObligInstRecMenorA30Dias '" & Format(pdFecha, "yyyymmdd") & "', '" & Format(pdFechaBalanceDiario, "yyyymmdd") & "'," & pnTipoCambBal & "," & pnTipoCambioMes & ""
    Set ObtenerObligInstRecMenora30Dias = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    Exit Function
ErrObtenerObligInstRecMenora30Dias:
    Call RaiseError(MyUnhandledError, "ObtenerObligInstRecMenora30Dias")
End Function '***NAGL Según ERS006-2019 20190430

Public Function ObtenerSaldoAdeudadoMyPCP(ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, ByVal psPlazo As String, ByVal psNacional As String) As Currency
On Error GoTo ObtenerSaldoAdeudadoMyPCPErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_MontoAdeudadosxPlazoyMonedaxCP '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "','" & psPlazo & "','" & psNacional & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!CortoCap), 0, oRs!CortoCap)
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSaldoAdeudadoMyPCP = nSaldo
Exit Function
ObtenerSaldoAdeudadoMyPCPErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerSaldoAdeudadoMyPCP Method")
End Function
'JIPR20200824 INICIO
Public Function MontoAdeudados15BCLP(ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, ByVal psPlazo As String) As Currency
On Error GoTo MontoAdeudados15BCLPErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_MontoAdeudados15BCLP '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "','" & psPlazo & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!CortoCap), 0, oRs!CortoCap)
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   MontoAdeudados15BCLP = nSaldo
Exit Function
MontoAdeudados15BCLPErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:MontoAdeudados15BCLP Method")
End Function
'JIPR20200824 FIN

Public Function ObtenerSaldoDiarioRestringido(pdFecDiaProceso As Date, ByVal psMoneda As String, ByVal pnPlazo As Integer) As Currency
On Error GoTo ObtenerSaldoDiarioRestringidoErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ObtenerRestringidos '" & Format(pdFecDiaProceso, "YYYYMMDD") & "','" & psMoneda & "'," & pnPlazo & ""
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
           oRs.MoveNext
      Loop
   Else
        nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSaldoDiarioRestringido = nSaldo
Exit Function
ObtenerSaldoDiarioRestringidoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerSaldoDiarioRestringido Method")
End Function
Public Function SaldoCtas(lnNroCol As Long, lsOpeCod As String, ByVal ldFecha As Date, ByVal pdFecFin As Date, ByVal pnTCF As Currency, ByVal pnTCFF As Currency) As Currency
Dim rsC As New ADODB.Recordset
Dim oRep As New DRepCtaColumna
Set rsC = oRep.GetRepColumnaCtaSaldo(lsOpeCod, lnNroCol, ldFecha, gbBitCentral)
SaldoCtas = 0
If Not RSVacio(rsC) Then
    If Mid(lsOpeCod, 3, 1) = "1" Then
        SaldoCtas = IIf(IsNull(rsC!Total), 0, rsC!Total)
    Else
        If lnNroCol = 29 Then
            SaldoCtas = Round(IIf(IsNull(rsC!TotalME), 0, rsC!TotalME), 1)
        Else
            If pdFecFin = ldFecha Then
               SaldoCtas = IIf(IsNull(rsC!Total), 0, Round(rsC!Total / pnTCFF, 2))
            Else
               SaldoCtas = IIf(IsNull(rsC!Total), 0, Round(rsC!Total / pnTCF, 2))
            End If
        End If
    End If
End If
Set oRep = Nothing
RSClose rsC
End Function
Public Function ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptaciones15B(ByVal pnNroCliente As String, ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, pnTipoCambio As Currency, Optional pnTipo As Integer = 0, Optional ByVal pnPlazo As Long = 99999) As ADODB.Recordset
On Error GoTo ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptaciones15BErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesX15B " & pnNroCliente & ",'" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'," & pnTipoCambio & "," & pnTipo & "," & pnPlazo
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptaciones15B = oRs
   oConec.CierraConexion
Exit Function
ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptaciones15BErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptaciones15B Method")
End Function
Public Function ObtenerSaldoDepositosGrandesAcreedores_Estado15B(ByVal psMoneda As String, ByVal pdFecDiaProceso As Date, ByVal pnTipoCambio As Currency, Optional ByVal pnPlazo As Long = 99999) As ADODB.Recordset
On Error GoTo ObtenerSaldoDepositosGrandesAcreedores_Estado15BErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SaldoDepositosGrandesAcreedores_Estadox15B '" & psMoneda & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & pnTipoCambio & "'," & pnPlazo
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerSaldoDepositosGrandesAcreedores_Estado15B = oRs
   oConec.CierraConexion
Exit Function
ObtenerSaldoDepositosGrandesAcreedores_Estado15BErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerSaldoDepositosGrandesAcreedores_Estado15B Method")
End Function

Public Function ObtenerHistEEFFMensual(ByVal dFecha As Date, ByVal psOpeCod As String, ByVal psMoneda As String) As ADODB.Recordset
On Error GoTo ObtenerHistEEFFMensualErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec NIIF_stp_sel_NIIF_HistEEFFMensual '" & Format(dFecha, "YYYY/MM/DD") & "','" & psOpeCod & "','" & psMoneda & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerHistEEFFMensual = oRs
   oConec.CierraConexion
Exit Function
ObtenerHistEEFFMensualErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerHistEEFFMensual Method")
End Function

Public Sub InsertarHistEEFFMensual(ByVal psMov As String, ByVal pdFechaReporte As Date, ByVal pnCorrelativo As Integer, ByVal pnMontoGenerado As Currency, ByVal psOpeCod As String, ByVal psMoneda As String)
On Error GoTo InsertarHistEEFFMensualErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec NIIF_stp_ins_NIIF_HistEEFFMensual '" & psMov & "','" & Format(pdFechaReporte, "YYYYMMDD") & "'," & pnCorrelativo & "," & pnMontoGenerado & ",'" & psOpeCod & "','" & psMoneda & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   oConec.CierraConexion
Exit Sub
InsertarHistEEFFMensualErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertarHistEEFFMensual Method")
End Sub
Public Function ObtenerHistEEFFMensualxCorrelativo(ByVal dFecha As Date, ByVal psOpeCod As String, ByVal psMoneda As String, ByVal pnCorrelativo As Integer) As Currency
On Error GoTo ObtenerHistEEFFMensualxCorrelativoErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   ObtenerHistEEFFMensualxCorrelativo = 0
   psSql = "exec NIIF_stp_sel_NIIF_HistEEFFMensualxCorrelativo '" & Format(dFecha, "YYYY/MM/DD") & "','" & psOpeCod & "','" & psMoneda & "'," & pnCorrelativo
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not (oRs.BOF Or oRs.EOF) Then
        ObtenerHistEEFFMensualxCorrelativo = IIf(IsNull(oRs!nMontoGenerado), 0, oRs!nMontoGenerado)
   End If
   oConec.CierraConexion
Exit Function
ObtenerHistEEFFMensualxCorrelativoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerHistEEFFMensualxCorrelativo Method")
End Function
Public Function SaldoAhoPlaFijCTS(ByVal ldFecha As String, ByVal nMoneda As Integer, ByVal cProducto As String) As Currency
Dim oNEncajeBCR As NEncajeBCR
Set oNEncajeBCR = New NEncajeBCR
Dim rsC As New ADODB.Recordset

Set rsC = oNEncajeBCR.ObtenerSaldoAhoPlaFijCTS(ldFecha, nMoneda, cProducto)
SaldoAhoPlaFijCTS = 0
    SaldoAhoPlaFijCTS = Round(rsC!nSaldo, 2)
Set oNEncajeBCR = Nothing
RSClose rsC
End Function
Public Function SaldoCracsAnexoDiario(ByVal ldFecha As String, ByVal nMoneda As Integer, ByVal cProducto As String) As Currency
Dim oNEncajeBCR As NEncajeBCR
Set oNEncajeBCR = New NEncajeBCR
Dim rsC As New ADODB.Recordset

Set rsC = oNEncajeBCR.ObtenerSaldoCracsEncajeDiario(ldFecha, nMoneda, cProducto)
SaldoCracsAnexoDiario = 0
If rsC.RecordCount > 0 Then
    SaldoCracsAnexoDiario = Round(rsC!nSaldo, 2)
Else
    SaldoCracsAnexoDiario = 0
End If

Set oNEncajeBCR = Nothing
RSClose rsC
End Function
Public Function SaldoCajasCracsAnexoDiario(ByVal ldFecha As String, ByVal nMoneda As Integer, ByVal cProducto As String) As Currency
Dim oNEncajeBCR As NEncajeBCR
Set oNEncajeBCR = New NEncajeBCR
Dim rsC As New ADODB.Recordset

Set rsC = oNEncajeBCR.ObtenerSaldoCajasCracsEncajeDiario(ldFecha, nMoneda, cProducto)
SaldoCajasCracsAnexoDiario = 0
If rsC.RecordCount > 0 Then
    SaldoCajasCracsAnexoDiario = Round(rsC!nSaldo, 2)
Else
    SaldoCajasCracsAnexoDiario = 0
End If
    
Set oNEncajeBCR = Nothing
RSClose rsC
End Function
Public Function SaldoChequeAhoPlaFijCTS(ByVal ldFecha As String, ByVal nMoneda As Integer, ByVal cProducto As String) As Currency
Dim oNEncajeBCR As NEncajeBCR
Set oNEncajeBCR = New NEncajeBCR
Dim rsC As New ADODB.Recordset

Set rsC = oNEncajeBCR.ObtenerSaldoChequeAhoPlaFijCTS(ldFecha, nMoneda, cProducto)
SaldoChequeAhoPlaFijCTS = 0
    SaldoChequeAhoPlaFijCTS = Round(rsC!nSaldo, 2)
Set oNEncajeBCR = Nothing
RSClose rsC
End Function
Public Function SaldoBCRPAnexoDiario(ByVal ldFecha As String, ByVal nMoneda As Integer) As Currency
Dim oNEncajeBCR As NEncajeBCR
Set oNEncajeBCR = New NEncajeBCR
Dim rsC As New ADODB.Recordset

Set rsC = oNEncajeBCR.ObtenerSaldoBCRPEncajeDiario(ldFecha, nMoneda)
SaldoBCRPAnexoDiario = 0
If rsC.RecordCount > 0 Then
    SaldoBCRPAnexoDiario = Round(rsC!nImporte, 2)
Else
    SaldoBCRPAnexoDiario = 0
End If

Set oNEncajeBCR = Nothing
RSClose rsC
End Function
Public Function SaldoCajasObligExoneradas(ByVal ldFecha As String, ByVal nMoneda As Integer) As Currency
Dim oNEncajeBCR As NEncajeBCR
Set oNEncajeBCR = New NEncajeBCR
Dim rsC As New ADODB.Recordset
Dim lnContador As Integer
Dim lbBol As Boolean
lnContador = 0
lbBol = False
Set rsC = oNEncajeBCR.ObtenerSaldoCajasObligExoneradas(ldFecha, nMoneda)
SaldoCajasObligExoneradas = 0
If rsC.BOF Or rsC.EOF Then
    lbBol = False
Else
    SaldoCajasObligExoneradas = Round(rsC!nMonto, 2)
    If SaldoCajasObligExoneradas <> 0 Then
        lbBol = True
    End If
End If
If lbBol = False Then
    Do While SaldoCajasObligExoneradas = 0 Or lnContador = 60
        lnContador = lnContador + 1
        ldFecha = Format(DateAdd("d", -1, Left(ldFecha, 4) + "/" + Mid(ldFecha, 5, 2) + "/" + Mid(ldFecha, 7, 2)), "YYYYMMDD")
        SaldoCajasObligExoneradas = SaldoCajasObligExoneradas(ldFecha, nMoneda)
    Loop
End If

Set oNEncajeBCR = Nothing
RSClose rsC
End Function
Public Function GetSaldoMEPosCambiaria(ByVal pdFecha As String, ByVal pcCtaCont As String, Optional pnTipo As String = "") As Currency
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
On Error GoTo GetSaldoMEPosCambiariaErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
    '****NAGL 20170803
    If pnTipo = "BitRep" Then
        sSQL = "exec stp_sel_SaldosMEPosCambiariaAnex15A '" & pdFecha & "','" & pcCtaCont & "'"
    Else
        sSQL = "exec stp_sel_SaldosMEPosCambiaria '" & pdFecha & "','" & pcCtaCont & "'"
    End If '**NAGL 20170803
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    If Not (rs.BOF Or rs.EOF) Then
      GetSaldoMEPosCambiaria = rs!nCtaSaldoImporteMEActual
    End If
    oCon.CierraConexion
    Exit Function
GetSaldoMEPosCambiariaErr:
    Call RaiseError(MyUnhandledError, "CTS")
End Function 'NAGL 20170803 ERS002-2017 Agregó pnTipo
Public Function ObtenerColocacSaldosDiarioTramos(ByVal pdFecha As Date, ByVal pnCodigo As Integer, ByVal psMoneda As String, Optional ByVal psConcepto As String = "") As Currency 'NAGL 20191014 Agregó psConcepto y Cambio en pnCodigo
On Error GoTo ObtenerColocacSaldosDiarioTramosErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nValor As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ColocacSaldosDiarioTramos '" & Format(pdFecha, "YYYYMMDD") & "'," & pnCodigo & ",'" & psConcepto & "'" 'NAGL 20191014 Agregó psConcepto y Cambio por pnCodigo
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        If psMoneda = "1" Then
          nValor = IIf(IsNull(oRs!nSaldoMN), 0, oRs!nSaldoMN)
        Else
          nValor = IIf(IsNull(oRs!nSaldoME), 0, oRs!nSaldoME)
        End If
        oRs.MoveNext
      Loop
   Else
    nValor = 0
   End If
   oConec.CierraConexion
   ObtenerColocacSaldosDiarioTramos = nValor
Exit Function
ObtenerColocacSaldosDiarioTramosErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerParamEncDiarioxCodigo Method")
End Function
'*******************************************************************************************
'ALPA20131204*******************************************************************************
Public Function ObtenerDatos15paraTxt(ByVal pdFecha As Date, ByVal psReporte As String) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta

On Error GoTo ObtenerDatos15paraTxtErr
    
    Set rs = New ADODB.Recordset
    sSQL = "exec stp_Reporte15Txt '" & Format(pdFecha, "YYYY/MM/DD") & "','" & psReporte & "'"

    oCon.AbreConexion
    Set ObtenerDatos15paraTxt = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Exit Function
ObtenerDatos15paraTxtErr:
    Call RaiseError(MyUnhandledError, "Datos para archivo txt de los reportes 15 A, B y C")
End Function

Public Function ObtenerOverNightTramosResidual(psMoneda As String, pdFecDiaProceso As Date, psTipo As String) As ADODB.Recordset
On Error GoTo ObtenerCtaContSaldoDiarioErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_OverNightTramosResidual '" & psMoneda & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psTipo & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerOverNightTramosResidual = oRs
   oConec.CierraConexion
Exit Function
ObtenerCtaContSaldoDiarioErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Function ObtenerCuentasxCobrarTramosResidual(psMoneda As String, pdFecDiaProceso As Date) As ADODB.Recordset
On Error GoTo ObtenerCuentasxCobrarTramosResidualErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_ERS0792016_sel_CuentasxCobrarORTramosResidual '" & psMoneda & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerCuentasxCobrarTramosResidual = oRs
   oConec.CierraConexion
Exit Function
ObtenerCuentasxCobrarTramosResidualErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function '****NAGL ERS 079-2016 20170407

Public Function ObtenerInversionesVentaTramosResidual(psMoneda As String, pdFecDiaProceso As Date) As ADODB.Recordset
On Error GoTo ObtenerInversionesVentaTramosResidualErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_ERS0792016_sel_InversionVentaResidual '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerInversionesVentaTramosResidual = oRs
   oConec.CierraConexion
Exit Function
ObtenerInversionesVentaTramosResidualErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function '****NAGL ERS 079-2016 20170407
Public Function ObtenerInversionesVenta16B(psMoneda As String, pdFecDiaProceso As Date, psCtaCod As String) As ADODB.Recordset
On Error GoTo ObtenerInversionesVenta16BErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_ERS0792016_sel_InversionVenta16B '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "','" & psCtaCod & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerInversionesVenta16B = oRs
   oConec.CierraConexion
Exit Function
ObtenerInversionesVenta16BErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InversionesVenta16B")
End Function '****NAGL ERS 079-2016 20170407

Public Function ObtenerInversionesAVencimientoResidual(pdFecDiaProceso As Date, pnMoneda As Integer) As ADODB.Recordset
On Error GoTo ObtenerInversionesAVencimientoResidualErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_InversionesAVencimiento_Residual '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'," & pnMoneda
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerInversionesAVencimientoResidual = oRs
   oConec.CierraConexion
Exit Function
ObtenerInversionesAVencimientoResidualErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Function ObtenerCreditosTramosResidual(psMoneda As String, pdFecDiaProceso As Date) As ADODB.Recordset
On Error GoTo ObtenerCreditosTramosResidualErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_CreditosTramosResidual '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerCreditosTramosResidual = oRs
   oConec.CierraConexion
Exit Function
ObtenerCreditosTramosResidualErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function

Public Function GetEstadSaldoCheques(ldFecha As Date, lsProducto As String, lsMoneda As String, Optional ByVal pnBitCentral As Boolean = False) As Currency
Dim TotalCheque As Currency
Dim prs As ADODB.Recordset
TotalCheque = 0

If pnBitCentral = True Then
 
    Select Case lsProducto
    
    Case "Plazo"
        psSql = "SELECT sum(nMonChqVal) as Monto "
        psSql = psSql & " FROM CAPESTADSALDO CS "
        psSql = psSql & " WHERE Convert(Varchar(8),CS.dEstad,112) LIKE '" & Format(ldFecha, "YYYYMMdd") & "%' "
        psSql = psSql & " AND NMONEDA=" & lsMoneda & " AND NPRODUCTO in (" & Producto.gCapPlazoFijo & ", " & Producto.gCapCTS & " ) "
        psSql = psSql & " and (nPersoneria <>4 and nPersoneria<>5 ) "
    
    Case "Ahorros"
        psSql = "SELECT sum(nMonChqVal) as Monto "
        psSql = psSql & " FROM CAPESTADSALDO CS "
        psSql = psSql & " WHERE Convert(Varchar(8),CS.dEstad,112) LIKE '" & Format(ldFecha, "YYYYMMdd") & "%' "
        psSql = psSql & " AND NMONEDA=" & lsMoneda & " AND NPRODUCTO in(" & Producto.gCapAhorros & ") "
        psSql = psSql & " and (nPersoneria <>4 and nPersoneria<>5 ) "
        
    End Select
    
Else

   Select Case lsProducto
       Case "Plazo"
               psSql = "SELECT SUM(nMonChqVal) AS Monto " _
                   & "FROM ESTADDIAPFCONSOL WHERE DATEDIFF(DAY,dEstadPF,'" & Format(ldFecha, gsFormatoFecha) & "')=0 AND cMoneda='" & lsMoneda & "' " _
                   & "Union " _
                   & "SELECT SUM(nMonChqVal) AS Monto " _
                   & "FROM ESTADDIACTSCONSOL WHERE DATEDIFF(DAY,dEstadcts,'" & Format(ldFecha, gsFormatoFecha) & "')=0 AND cMoneda='" & lsMoneda & "' "
           
       Case "Ahorros"
               psSql = "SELECT dEstadAC, SUM(nMonChqVal+nChqCMAC) as Monto " _
                   & "FROM ESTADDIAACCONSOL WHERE DATEDIFF(DAY,dEstadAC,'" & Format(ldFecha, gsFormatoFecha) & "')=0 AND cMoneda='" & lsMoneda & "' " _
                   & "GROUP BY  dEstadAC"
   End Select
End If
Set prs = dbConec.CargaRecordSet(psSql)
Do While Not prs.EOF
   TotalCheque = TotalCheque + IIf(IsNull(prs!Monto), 0, prs!Monto)
   prs.MoveNext
Loop
prs.Close
GetEstadSaldoCheques = TotalCheque
End Function

Public Function ObtenerCtaContSaldoBalanceDiarioxLike(psCtaContCod As String, pdFecDiaProceso As Date, psMoneda As String, pnTipoCambio As Currency) As Currency
On Error GoTo ObtenerCtaContSaldoBalanceDiarioxLikeErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ObtenerDatosBalanceDiarioxLike '" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'," & pnTipoCambio
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = oRs!nSaldoFinImporte
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerCtaContSaldoBalanceDiarioxLike = nSaldo
Exit Function
ObtenerCtaContSaldoBalanceDiarioxLikeErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerCtaContSaldoBalanceDiarioxLike Method")
End Function

Public Function ObtenerInteresDevengadoResidual(pdFecDiaProceso As Date, psMoneda As String) As ADODB.Recordset
On Error GoTo ObtenerInteresDevengadoResidualErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ObtnerInteresDevengadoResidual '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerInteresDevengadoResidual = oRs
   oConec.CierraConexion
Exit Function
ObtenerInteresDevengadoResidualErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerInteresDevengadoResidual: Method")
End Function

Public Function ObtenerInteresFuturoResidual(pdFecha As Date, psTipoCambio As String, psEstado As String) As ADODB.Recordset
On Error GoTo ObtenerInteresFuturoResidualErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   If psEstado = "VIG" Then
        psSql = "exec stp_sel_ObtenerInteresFuturoResidualVIG '" & Format(pdFecha, "yyyymmdd") & "','" & psTipoCambio & "'"
   Else 'REF
        psSql = "exec stp_sel_ObtenerInteresFuturoResidualREF '" & Format(pdFecha, "yyyymmdd") & "','" & psTipoCambio & "'"
   End If
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerInteresFuturoResidual = oRs
   oConec.CierraConexion
Exit Function
ObtenerInteresFuturoResidualErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerInteresFuturoResidual: Method")
End Function '***NAGL 20191030 Según Anx03_ERS006-2019

'FRHU 20131205, RQ13649
Public Function recuperarConfParametroUtilidad(ByVal pnAnio As Integer) As ADODB.Recordset
On Error GoTo recuperarConfParametroUtilidadError
   Dim oDConecta As DConecta
   Set oDConecta = New DConecta
   
   oDConecta.AbreConexion
   psSql = "exec stp_sel_RecuperarConfParametroUtilidad " & pnAnio
   Set recuperarConfParametroUtilidad = oDConecta.CargaRecordSet(psSql)
   oDConecta.CierraConexion
Exit Function
recuperarConfParametroUtilidadError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:recuperarConfParametroUtilidad Method")
End Function
Public Sub ModificarInsertarConfUtilidad(ByVal pnParamVar As Integer, ByVal pnAnio As Integer, ByVal pnValor As Integer, ByVal pnAnioFE As Integer)
On Error GoTo ModificarInsertarConfUtilidadError
    Dim oDConecta As DConecta
    Set oDConecta = New DConecta
    oDConecta.AbreConexion
    psSql = "stp_ins_upd_ConfUtilidad " & pnParamVar & ", " & pnAnio & ", " & pnValor & ", " & pnAnioFE
    oDConecta.Ejecutar (psSql)
    oDConecta.CierraConexion
    Exit Sub
ModificarInsertarConfUtilidadError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ModificarInsertarConfUtilidad Method")
End Sub
Public Function recuperarReporteAnexo01(ByVal pcAnio As String, ByVal pnSemestre As Integer, ByVal pnTipo As Integer) As ADODB.Recordset
On Error GoTo recuperarReporteAnexo01Error
   Dim oDConecta As DConecta
   Set oDConecta = New DConecta
   
   oDConecta.AbreConexion
   psSql = "exec stp_sel_ReporteAnexo01PropuDistribuUtilidades '" & pcAnio & "', " & pnSemestre & ", " & pnTipo
   Set recuperarReporteAnexo01 = oDConecta.CargaRecordSet(psSql)
   oDConecta.CierraConexion
Exit Function
recuperarReporteAnexo01Error:
   Call RaiseError(MyUnhandledError, "DBalanceCont:recuperarReporteAnexo01 Method")
End Function
'FIN FRHU 20131205
'FRHU 20131209, RQ13650
Public Function recuperarReporteAnexo02(ByVal pcAnio As String, ByVal pnSemestre As Integer, ByVal pnTipo As Integer) As ADODB.Recordset
On Error GoTo recuperarReporteAnexo02Error
   Dim oDConecta As DConecta
   Set oDConecta = New DConecta
   
   oDConecta.AbreConexion
   psSql = "exec stp_sel_ReporteAnexo02PropuCapitaliUtilidades '" & pcAnio & "', " & pnSemestre & ", " & pnTipo
   Set recuperarReporteAnexo02 = oDConecta.CargaRecordSet(psSql)
   oDConecta.CierraConexion
Exit Function
recuperarReporteAnexo02Error:
   Call RaiseError(MyUnhandledError, "DBalanceCont:recuperarReporteAnexo02 Method")
End Function
'FIN FRHU 20131209
'FRHU 20131219 RQ13656
Public Function recuperarEstadoDeCambioPatrimonio(ByVal pcAnio As String, ByVal pnSemestre As Integer) As ADODB.Recordset
On Error GoTo recuperarEstadoDeCambioPatrimonioError
   Dim oDConecta As DConecta
   Set oDConecta = New DConecta
   
   oDConecta.AbreConexion
   psSql = "exec stp_sel_RepEstadoDeCambioEnElPatrimonio '" & pcAnio & "', " & pnSemestre
   Set recuperarEstadoDeCambioPatrimonio = oDConecta.CargaRecordSet(psSql)
   oDConecta.CierraConexion
Exit Function
recuperarEstadoDeCambioPatrimonioError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:recuperarEstadoDeCambioPatrimonio Method")
End Function
'FIN FRHU 20131219 RQ13656

' ACTIVOS FIJOS

'FRHU 20131209 RQ13651
Public Function RepAFDepreciacionAcumuladaIntangible(ByVal pcAnio As String, ByVal pnSemestre As String) As ADODB.Recordset
On Error GoTo RepAFDepreciacionAcumuladaIntangibleError
   Dim oDConecta As DConecta
   Set oDConecta = New DConecta
   
   oDConecta.AbreConexion
   psSql = "exec stp_sel_ActivoFijoDepreciacionAcumuladaIntangible '" & pcAnio & "', '" & pnSemestre & "'"
   Set RepAFDepreciacionAcumuladaIntangible = oDConecta.CargaRecordSet(psSql)
   oDConecta.CierraConexion
Exit Function
RepAFDepreciacionAcumuladaIntangibleError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:RepAFDepreciacionAcumuladaIntangible Method")
End Function
'FIN FRHU 20131211 RQ13652

'FRHU 20131214 RQ13652
Public Function RepAFActivoFijoIntangible(ByVal pcAnio As String, ByVal pnSemestre As String) As ADODB.Recordset
On Error GoTo RepAFActivoFijoIntangibleError
   Dim oDConecta As DConecta
   Set oDConecta = New DConecta
   
   oDConecta.AbreConexion
   psSql = "exec stp_sel_ActivoFijoEIntangible '" & pcAnio & "', '" & pnSemestre & "'"
   Set RepAFActivoFijoIntangible = oDConecta.CargaRecordSet(psSql)
   oDConecta.CierraConexion
Exit Function
RepAFActivoFijoIntangibleError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:RepAFActivoFijoIntangible Method")
End Function
'FIN FRHU 20131214 RQ13652

'FRHU 20131216 RQ13653
Public Function RepAFControlIntangible(ByVal pcAnio As String, ByVal pnSemestre As Integer) As ADODB.Recordset
On Error GoTo RepAFControlIntangibleError
   Dim oDConecta As DConecta
   Set oDConecta = New DConecta
   
   oDConecta.AbreConexion
   psSql = "exec stp_sel_ActivoFijoControlIntangibleOtrosActivos '" & pcAnio & "', " & pnSemestre
   Set RepAFControlIntangible = oDConecta.CargaRecordSet(psSql)
   oDConecta.CierraConexion
Exit Function
RepAFControlIntangibleError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:RepAFControlIntangible Method")
End Function
'FIN FRHU 20131216 RQ13653
'FRHU 20131217 RQ13654
Public Function RepAFControlInmuebleMaquinaria(ByVal pcAnio As String, ByVal pnSemestre As Integer) As ADODB.Recordset
On Error GoTo RepAFControlInmuebleMaquinariaError
   Dim oDConecta As DConecta
   Set oDConecta = New DConecta
   
   oDConecta.AbreConexion
   psSql = "exec stp_sel_ActivoFijoControlInmuebleMaquinariaEquipo '" & pcAnio & "', " & pnSemestre
   Set RepAFControlInmuebleMaquinaria = oDConecta.CargaRecordSet(psSql)
   oDConecta.CierraConexion
Exit Function
RepAFControlInmuebleMaquinariaError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:RepAFControlInmuebleMaquinaria Method")
End Function
'FIN FRHU 20131217 RQ13654

'FRHU 20131218 RQ13655
Public Function RepAFControlDepreciacion(ByVal pcAnio As String, ByVal pnSemestre As Integer) As ADODB.Recordset
On Error GoTo RepAFControlDepreciacionError
   Dim oDConecta As DConecta
   Set oDConecta = New DConecta
   
   oDConecta.AbreConexion
   psSql = "exec stp_sel_ActivoFijoControlDeDepreciacion '" & pcAnio & "', " & pnSemestre
   Set RepAFControlDepreciacion = oDConecta.CargaRecordSet(psSql)
   oDConecta.CierraConexion
Exit Function
RepAFControlDepreciacionError:
   Call RaiseError(MyUnhandledError, "DBalanceCont:RepAFControlDepreciacion Method")
End Function
'FIN FRHU 20131218 RQ13655
'ALPA20140114*************
Public Function ObtenerCreditosRTramosResidual(psMoneda As String, pdFecDiaProceso As Date) As ADODB.Recordset
On Error GoTo ObtenerCreditosRTramosResidualErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_CreditosRTramosResidual '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerCreditosRTramosResidual = oRs
   oConec.CierraConexion
Exit Function
ObtenerCreditosRTramosResidualErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerCreditosRTramosResidual Method")
End Function
'*************************
'ALPA20140114*************
Public Function ObtenerFondeoEncajexTramos(ByVal pdFecha As Date, ByVal pnTipoCambio As Currency) As ADODB.Recordset
    Dim lsSql As String
    Dim oCon As New DConecta
    Dim oRs As New ADODB.Recordset
    Dim lnSuma As Currency
On Error GoTo ErrObtenerFondeoEncajexTramos
    oCon.AbreConexion
    lnSuma = 0
    lsSql = "Exec stp_sel_FondeoEncajexTramos '" & Format(pdFecha, "yyyymmdd") & "'," & pnTipoCambio & ""
    oCon.AbreConexion
    Set ObtenerFondeoEncajexTramos = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    Exit Function
ErrObtenerFondeoEncajexTramos:
    Call RaiseError(MyUnhandledError, "ObtenerFondeoEncajexTramos")
End Function
Public Function ObtenerSaldoDepositosGrandesAcreedores_EstadoxTramos(ByVal psMoneda As String, ByVal pdFecDiaProceso As Date, ByVal pnTipoCambio As Currency) As ADODB.Recordset
On Error GoTo ObtenerSaldoDepositosGrandesAcreedores_EstadoxTramosErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SaldoDepositosGrandesAcreedores_EstadoxTramos '" & psMoneda & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & pnTipoCambio & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerSaldoDepositosGrandesAcreedores_EstadoxTramos = oRs
   oConec.CierraConexion
Exit Function
ObtenerSaldoDepositosGrandesAcreedores_EstadoxTramosErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerSaldoDepositosGrandesAcreedores_EstadoxTramos Method")
End Function
Public Function ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramos(ByVal pnNroCliente As String, ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, pnTipoCambio As Currency, Optional pnTipo As Integer = 0) As ADODB.Recordset
On Error GoTo ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramosErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramos " & pnNroCliente & ",'" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'," & pnTipoCambio & "," & pnTipo
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramos = oRs
   oConec.CierraConexion
Exit Function
ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramosErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramos Method")
End Function
Public Function ObtenerFondeoEncajexTramosxProducto(ByVal pdFecha As Date, ByVal pnTipoCambio As Currency, ByVal psProducto As String) As ADODB.Recordset
    Dim lsSql As String
    Dim oCon As New DConecta
    Dim oRs As New ADODB.Recordset
    Dim lnSuma As Currency
On Error GoTo ErrObtenerFondeoEncajexTramosxProducto
    oCon.AbreConexion
    lnSuma = 0
    lsSql = "Exec stp_sel_FondeoEncajexTramosxProducto '" & Format(pdFecha, "yyyymmdd") & "'," & pnTipoCambio & ",'" & psProducto & "'"
    oCon.AbreConexion
    Set ObtenerFondeoEncajexTramosxProducto = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    Exit Function
ErrObtenerFondeoEncajexTramosxProducto:
    Call RaiseError(MyUnhandledError, "ObtenerFondeoEncajexTramosxProducto")
End Function
'NAGL 20190513 Cambió de ObtenerFondeoGirosxPagar a ObtenerFondeoObligVista
Public Function ObtenerFondeoObligVista(ByVal pdFecha As Date, ByVal pnTipoCambio As Currency) As ADODB.Recordset
Dim lsSql As String
Dim oCon As New DConecta
Dim oRs As New ADODB.Recordset
On Error GoTo ErrObtenerFondeoObligVista
    oCon.AbreConexion
    'NAGL 20190513 Cambió de stp_ERS0792016_sel_FondeoGirosxPagar a stp_ERS0062017_sel_Fondeo_ObligVista
    lsSql = "Exec stp_ERS0062017_sel_Fondeo_ObligVista '" & Format(pdFecha, "yyyymmdd") & "'," & pnTipoCambio & ""
    oCon.AbreConexion
    Set ObtenerFondeoObligVista = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    Exit Function
ErrObtenerFondeoObligVista:
    Call RaiseError(MyUnhandledError, "ObtenerFondeoObligVista")
End Function '****NAGL ERS 079-2016 2017
Public Function ObtenerFondeoRetencionesJudiciales(ByVal pdFecha As Date, ByVal pnTipoCambio As Currency) As ADODB.Recordset
Dim lsSql As String
Dim oCon As New DConecta
Dim oRs As New ADODB.Recordset
On Error GoTo ErrObtenerFondeoRetencionesJudiciales
    oCon.AbreConexion
    lsSql = "Exec stp_ERS0792016_sel_FondeoRetencionesJudiciales '" & Format(pdFecha, "yyyymmdd") & "'," & pnTipoCambio & ""
    oCon.AbreConexion
    Set ObtenerFondeoRetencionesJudiciales = oCon.CargaRecordSet(lsSql)
    oCon.CierraConexion
    Exit Function
ErrObtenerFondeoRetencionesJudiciales:
    Call RaiseError(MyUnhandledError, "ObtenerFondeoRetencionesJudiciales")
End Function '****NAGL ERS 079-2016 2017

Public Function ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramosxProducto(ByVal pnNroCliente As String, ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, pnTipoCambio As Currency, Optional pnTipo As Integer = 0, Optional psProducto As String = "233") As ADODB.Recordset
On Error GoTo ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramosxProductoErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramosxProducto " & pnNroCliente & ",'" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'," & pnTipoCambio & "," & pnTipo & ",'" & psProducto & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramosxProducto = oRs
   oConec.CierraConexion
Exit Function
ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramosxProductoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerSaldoDepositosGrandesAcreedores_Estado_DevolverClientesMayoresCaptacionesxTramosxProducto Method")
End Function
Public Function ObtenerSaldoDepositosGrandesAcreedores_EstadoxTramosxProducto(ByVal psMoneda As String, ByVal pdFecDiaProceso As Date, ByVal pnTipoCambio As Currency, Optional psProducto As String = "233") As ADODB.Recordset
On Error GoTo ObtenerSaldoDepositosGrandesAcreedores_EstadoxTramosxProductoErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SaldoDepositosGrandesAcreedores_EstadoxTramosxProducto '" & psMoneda & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & pnTipoCambio & "','" & psProducto & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerSaldoDepositosGrandesAcreedores_EstadoxTramosxProducto = oRs
   oConec.CierraConexion
Exit Function
ObtenerSaldoDepositosGrandesAcreedores_EstadoxTramosxProductoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerSaldoDepositosGrandesAcreedores_EstadoxTramosxProducto Method")
End Function

Public Function ObtenerRestringidosxTramosxProducto(ByVal psMoneda As String, ByVal pdFecDiaProceso As Date, ByVal pnPlazo As Integer, Optional psProducto As String = "233") As ADODB.Recordset
On Error GoTo ObtenerRestringidosxTramosxProductoErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ObtenerRestringidosxTramosxProducto '" & Format(pdFecDiaProceso, "YYYYMMDD") & "','" & psMoneda & "'," & pnPlazo & ",'" & psProducto & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerRestringidosxTramosxProducto = oRs
   oConec.CierraConexion
Exit Function
ObtenerRestringidosxTramosxProductoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerRestringidosxTramosxProducto Method")
End Function

Public Function ObtenerDepositosSistemaFinancieroyOFIxTramosxOProducto(ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, ByVal psProducto As String) As ADODB.Recordset
On Error GoTo ObtenerDepositosSistemaFinancieroyOFIxTramosxOProductoErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_DepositosSistemaFinancieroyOFIxTramosxOProducto '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "','" & psProducto & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerDepositosSistemaFinancieroyOFIxTramosxOProducto = oRs
   oConec.CierraConexion
Exit Function
ObtenerDepositosSistemaFinancieroyOFIxTramosxOProductoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerDepositosSistemaFinancieroyOFIxTramosxOProducto Method")
End Function

Public Function ObtenerMontoAdeudadosxPlazoyMonedaxCPxTramos(ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, ByVal psNacional As String, Optional pnTipoCambio As Currency = 0) As ADODB.Recordset 'NAGL 202012 Agregó pnTipoCambio
On Error GoTo ObtenerMontoAdeudadosxPlazoyMonedaxCPxTramosErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_MontoAdeudadosxPlazoyMonedaxCPxTramos '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "','" & psNacional & "'," & pnTipoCambio & "" 'NAGL 202012 Agregó pnTipoCambio
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerMontoAdeudadosxPlazoyMonedaxCPxTramos = oRs
   oConec.CierraConexion
Exit Function
ObtenerMontoAdeudadosxPlazoyMonedaxCPxTramosErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerDepositosSistemaFinancieroyOFIxTramosxOProducto Method")
End Function

Public Function ObtenerAdeudadosExterior(ByVal pdFecDiaProceso As Date, ByVal psMoneda As String, ByVal psMacional As String) As Currency
On Error GoTo ObtenerAdeudadosExteriorErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_ERS0792016_sel_AdeudadosExterior '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "','" & psMacional & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
        oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
Exit Function
ObtenerAdeudadosExteriorErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerAdeudadosExterior Method")
End Function '****NAGL ERS 079-2016 20170407

Public Function ObtenerIndicadores16Acredores(ByVal pdFecDiaProceso As Date, ByVal pnPlazo As Integer, ByVal pnTipoCambio As Currency, ByVal pnTop As Integer, ByVal pnTipo As Integer) As Currency
On Error GoTo ObtenerIndicadores16AcredoresErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   ObtenerIndicadores16Acredores = 0
   oConec.AbreConexion
   psSql = "exec stp_sel_Indicadores16Acredores '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'," & pnPlazo & "," & pnTipoCambio & "," & pnTop & "," & pnTipo
   ObtenerIndicadores16Acredores = 0
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not (oRs.BOF Or oRs.EOF) Then
        Do While Not oRs.EOF
        ObtenerIndicadores16Acredores = ObtenerIndicadores16Acredores + oRs!nSaldo
        oRs.MoveNext
   Loop
   End If
   
   oConec.CierraConexion
Exit Function
ObtenerIndicadores16AcredoresErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerIndicadores16Acredores Method")
End Function

Public Function ObtenerIndicadores16Depositantes(ByVal pdFecDiaProceso As Date, ByVal pnPlazo As Integer, ByVal pnTipoCambio As Currency, ByVal pnTop As Integer, ByVal pnTipo As Integer) As Currency
On Error GoTo ObtenerIndicadores16DepositantesErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   ObtenerIndicadores16Depositantes = 0
   psSql = "exec stp_sel_Indicadores16Acredores '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'," & pnPlazo & "," & pnTipoCambio & "," & pnTop & "," & pnTipo
   ObtenerIndicadores16Depositantes = 0
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not (oRs.BOF Or oRs.EOF) Then
   Do While Not oRs.EOF
        ObtenerIndicadores16Depositantes = ObtenerIndicadores16Depositantes + oRs!nSaldo
        oRs.MoveNext
   Loop
   End If
   
   oConec.CierraConexion
Exit Function
ObtenerIndicadores16DepositantesErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerIndicadores16Depositantes Method")
End Function

Public Function ObtenerDetalleAcreedoresDepositantes(ByVal pdFecDiaProceso As Date, ByVal pnTipoCambio As Currency, ByVal pnTop As Integer, ByVal pnTipo As Integer) As ADODB.Recordset
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oCon As New DConecta
  
On Error GoTo ObtenerDetalleAcreedoresDepositantesErr
    oCon.AbreConexion
    Set rs = New ADODB.Recordset
    sSQL = "exec stp_ERS0792016_sel_Indicadores16Acreedores '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'," & pnTipoCambio & "," & pnTop & "," & pnTipo
    oCon.AbreConexion
    Set rs = oCon.CargaRecordSet(sSQL)
    oCon.CierraConexion
    Set ObtenerDetalleAcreedoresDepositantes = rs
    Exit Function
ObtenerDetalleAcreedoresDepositantesErr:
    Call RaiseError(MyUnhandledError, "Detalle de Depositantes y Acreedores")
End Function 'NAGL ERS 079-2016 20170407

Public Function ObtenerIndicadores16FondeoCubierto(ByVal pdFecDiaProceso As Date, ByVal pnTipoCambio As Currency) As Currency
On Error GoTo ObtenerIndicadores16FondeoCubiertoErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_FondeoFondoCubierto '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'," & pnTipoCambio & ""
   ObtenerIndicadores16FondeoCubierto = 0
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not (oRs.BOF Or oRs.EOF) Then
        ObtenerIndicadores16FondeoCubierto = oRs!nSaldo
   End If
   oConec.CierraConexion
Exit Function
ObtenerIndicadores16FondeoCubiertoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerIndicadores16FondeoCubierto Method")
End Function

Public Function ObtenerFondeoFSD(ByVal pdFecDiaProceso As Date, ByVal pnTipoCambio As Currency, ByVal pnPlazo As Integer, ByVal pnCubierto As Integer) As Currency
On Error GoTo ObtenerFondeoFSDErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_ERS0792016_sel_FondeoFSD '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'," & pnTipoCambio & "," & pnPlazo & "," & pnCubierto & ""
   ObtenerFondeoFSD = 0
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not (oRs.BOF Or oRs.EOF) Then
        ObtenerFondeoFSD = oRs!nSaldo
   End If
   oConec.CierraConexion
Exit Function
ObtenerFondeoFSDErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerFondeoFSD Method")
End Function '*****NAGL ERS 079-2017 20170407

Public Function ObtenerSaldoDepositosGrandesAcreedores_EstadoxPlazo(ByVal psMoneda As String, ByVal pdFecDiaProceso As Date, ByVal pnTipoCambio As Currency, ByVal pnPlazo As Integer) As Currency
On Error GoTo ObtenerSaldoDepositosGrandesAcreedores_EstadoxPlazoErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SaldoDepositosGrandesAcreedores_EstadoxPlazo '" & psMoneda & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & pnTipoCambio & "'," & pnPlazo & ""
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
        oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerSaldoDepositosGrandesAcreedores_EstadoxPlazo = nSaldo
Exit Function
ObtenerSaldoDepositosGrandesAcreedores_EstadoxPlazoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerSaldoDepositosGrandesAcreedores_EstadoxPlazo Method")
End Function

Public Function ObtenerCtaContSaldoBalanceDiarioFindeMes(psCtaContCod As String, pdFecDiaProceso As Date, psMoneda As String, pnTipoCambio As Currency) As Currency
On Error GoTo ObtenerCtaContSaldoBalanceDiarioFindeMesErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ObtenerDatosBalanceDiarioFindeMes '" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'," & pnTipoCambio
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = Round(oRs!nSaldoFinImporte, 2)
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerCtaContSaldoBalanceDiarioFindeMes = nSaldo
Exit Function
ObtenerCtaContSaldoBalanceDiarioFindeMesErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function
'ALPA 20140901
Public Function obtenerSumaConsolidadoCtasxConcentracionFondosxPlazo(ByVal pdFecha As Date, ByVal pnTpoMoneda As Integer, ByVal psListaTpoIfi As String, ByVal psListaCodPersIFiExcluidas As String, Optional pnPlazo As Long = 99999) As Currency
    Dim lsSql As String
    Dim oCon As New DConecta
    Dim oRs As New ADODB.Recordset
    Dim lnSuma As Currency
    Dim lsCtaIFCodCtaCorriente As String, lsCtaIFCodCtaAhorro As String, lsCtaIFCodOverNightCod As String
    Dim lsCtaIFDescCtaCorriente As String, lsCtaIFDescCtaAhorro As String
On Error GoTo ErrobtenerSumaConsolidadoCtasxConcentracionFondosxPlazo
    oCon.AbreConexion
    lnSuma = 0
    lsSql = "Exec stp_sel_ListaConsolCtaRptConcentracionxPlazo '" & Format(pdFecha, "yyyymmdd") & "','" & pnTpoMoneda & "','" & psListaTpoIfi & "','" & psListaCodPersIFiExcluidas & "'," & pnPlazo
    oCon.AbreConexion
    Set oRs = oCon.CargaRecordSet(lsSql)
    If Not (oRs.BOF Or oRs.EOF) Then
        Do While Not oRs.EOF
        'lnSuma = lnSuma + IIf(IsNull(oRs!nSaldoCtaCorriente), 0, oRs!nSaldoCtaCorriente) + IIf(IsNull(oRs!nSaldoCtaAhorro), 0, oRs!nSaldoCtaAhorro) + IIf(IsNull(oRs!nSaldoDPFOVer), 0, oRs!nSaldoDPFOVer)
        If oRs!cPerscod <> "1090100822183" Then
            'If lsCtaIFCodCtaCorriente <> oRs!cCtaIFCodCtaCorriente And oRs!cCtaIFDescCtaCorriente <> oRs!cCtaIFDescCtaCorriente Then
            If lsCtaIFDescCtaCorriente <> oRs!cCtaIFDescCtaCorriente Then
                lnSuma = lnSuma + IIf(IsNull(oRs!nSaldoCtaCorriente), 0, oRs!nSaldoCtaCorriente)
            End If
            'If lsCtaIFCodCtaAhorro <> oRs!cCtaIFCodCtaAhorro And lsCtaIFDescCtaAhorro <> oRs!cCtaIFDescCtaAhorro Then
            If lsCtaIFCodCtaAhorro <> oRs!cCtaIFCodCtaAhorro Then
                lnSuma = lnSuma + IIf(IsNull(oRs!nSaldoCtaAhorro), 0, oRs!nSaldoCtaAhorro)
            End If
            If lsCtaIFCodOverNightCod <> oRs!cCtaIFCodOverNightCod Then
                lnSuma = lnSuma + IIf(oRs!cCtaIFCodOverNight = "PF", IIf(IsNull(oRs!nSaldoDPFOVer), 0, oRs!nSaldoDPFOVer), 0)
            End If
            'lnSuma = lnSuma + IIf(IsNull(oRs!nSaldoCtaCorriente), 0, oRs!nSaldoCtaCorriente) + IIf(IsNull(oRs!nSaldoCtaAhorro), 0, oRs!nSaldoCtaAhorro) + IIf(oRs!cCtaIFCodOverNight = "PF", IIf(IsNull(oRs!nSaldoDPFOVer), 0, oRs!nSaldoDPFOVer), 0)
            
            lsCtaIFCodCtaCorriente = oRs!cCtaIFCodCtaCorriente
            lsCtaIFCodCtaAhorro = oRs!cCtaIFCodCtaAhorro
            lsCtaIFCodOverNightCod = oRs!cCtaIFCodOverNightCod
            lsCtaIFDescCtaCorriente = oRs!cCtaIFDescCtaCorriente
            lsCtaIFDescCtaAhorro = oRs!cCtaIFDescCtaAhorro
        End If
        oRs.MoveNext
        Loop
    End If
    oCon.CierraConexion
    obtenerSumaConsolidadoCtasxConcentracionFondosxPlazo = lnSuma
    Exit Function
ErrobtenerSumaConsolidadoCtasxConcentracionFondosxPlazo:
    Call RaiseError(MyUnhandledError, "Recupera Ctas Concentración")
End Function

Public Function ObtenerDispobiblesenSFNxPlazo(psPersCod As String, psCtaContCod As String, pdFecDiaProceso As Date, nTipo As Integer, Optional nPlazo As Long = 99999) As Currency
On Error GoTo ObtenerDispobiblesenSFNxPlazoErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_DispobiblesenSFNxPlazo '" & psPersCod & "','" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'," & nTipo & "," & nPlazo
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
           oRs.MoveNext
      Loop
   Else
        nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerDispobiblesenSFNxPlazo = nSaldo
Exit Function
ObtenerDispobiblesenSFNxPlazoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBalanceSectorial Method")
End Function
Public Function RecuperarPorcentajeGastosPorAgencia(ByVal psAgeCod As String) As ADODB.Recordset
    Dim oDConecta As New DConecta
    On Error GoTo ErrRecuperarPorcentajeGastosPorAgencia
    oDConecta.AbreConexion
    psSql = "exec stp_sel_AgenciaPorcentajeGastosxAgencia '" & psAgeCod & "'"
    Set RecuperarPorcentajeGastosPorAgencia = oDConecta.CargaRecordSet(psSql)
    oDConecta.CierraConexion
    Set oDConecta = Nothing
    Exit Function
ErrRecuperarPorcentajeGastosPorAgencia:
   Call RaiseError(MyUnhandledError, "DBalanceCont:RecuperarPorcentajeGastosPorAgencia Method")
End Function
'Public Sub ActualizarEmpleadosPorcentajeCantidadxAgencia(ByVal psPeriodo As String, ByVal psCodAge As String)
'Dim rs As ADODB.Recordset
'Dim oConec As DConecta
'Set oConec = New DConecta
'Dim sql As String
'On Error GoTo ErrorActualizarEmpleadosPorcentajeCantidadxAgencia
'
'    sql = "exec stp_sel_EmpleadosCandidadPorcentajexAgencia '" & psPeriodo & "' , '" & psCodAge & "'"
'    oConec.AbreConexion
'    oConec.CargaRecordSet (sql)
'    oConec.CierraConexion
'    Exit Sub
'ErrorActualizarEmpleadosPorcentajeCantidadxAgencia:
'    Err.Raise Err.Number, "ActualizarEmpleadosPorcentajeCantidadxAgencia", Err.Description
'End Sub
'ALPA 20141219****************************************************************************
Public Function ObtenerCuentasInactivasRestringidas(ByVal pdFechaProceso As Date) As ADODB.Recordset
On Error GoTo ObtenerCuentasInactivasRestringidasErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ReporteCuentasInactivasRestringidas '" & Format(pdFechaProceso, "YYYY/MM/DD") & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerCuentasInactivasRestringidas = oRs
   oConec.CierraConexion
Exit Function
ObtenerCuentasInactivasRestringidasErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerCuentasInactivasRestringidasErr Method")
End Function
'*****************************************************************************************
'ALPA 20150129****************************************************************************
Public Function ObtenerEmpleadosPorcentajeCantidadxAgencia(ByVal psPeriodo As String, ByVal psCodAge As String) As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim oConec As DConecta
Set oConec = New DConecta
Dim sql As String
On Error GoTo ErrorObtenerEmpleadosPorcentajeCantidadxAgencia
    
    sql = "exec stp_sel_EmpleadosCandidadPorcentajexAgencia '" & psPeriodo & "' , '" & psCodAge & "'"
    oConec.AbreConexion
    Set ObtenerEmpleadosPorcentajeCantidadxAgencia = oConec.CargaRecordSet(sql)
    oConec.CierraConexion
    Exit Function
ErrorObtenerEmpleadosPorcentajeCantidadxAgencia:
    Err.Raise Err.Number, "ObtenerEmpleadosPorcentajeCantidadxAgencia", Err.Description
End Function
'*****************************************************************************************
'ALPA 20150520****************************************************************************
Public Function ObtenerCtaContSaldoCapitalDeveg(pdFecDiaProceso As Date, pnTipoCambio As Currency, psConcepto As String) As ADODB.Recordset
On Error GoTo ObtenerCtaContSaldoCapitalDevegErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SaldoCapitalDevengCreditosxPlazo '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'," & pnTipoCambio & ",'" & psConcepto & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerCtaContSaldoCapitalDeveg = oRs
   oConec.CierraConexion
Exit Function
ObtenerCtaContSaldoCapitalDevegErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerCtaContSaldoCapitalDeveg Method")
End Function
'*******************************************************************************************
'ALPA 20150520****************************************************************************
Public Function ObtenerCtaSaldoDiferidosCreditosxPlazo(pdFecDiaProceso As Date, pnTipoCambio As Currency) As ADODB.Recordset
On Error GoTo ObtenerCtaSaldoDiferidosCreditosxPlazoErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_SaldoDiferidosCreditosxPlazo '" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "'," & pnTipoCambio & ""
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerCtaSaldoDiferidosCreditosxPlazo = oRs
   oConec.CierraConexion
Exit Function
ObtenerCtaSaldoDiferidosCreditosxPlazoErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerCtaSaldoDiferidosCreditosxPlazo Method")
End Function
'*******************************************************************************************
'PASIERS0242015*******************************************************************
Public Function ObtieneRepSBS13InversionesCapitalSocialSubsidiarias(ByVal psAnio As String, ByVal psMes As String) As Double
    Dim rs As ADODB.Recordset
On Error GoTo ERRORObtiene
    psSql = "stp_sel_ERS0242015_ObtieneRepSBS13InversionesCapitalSocialSubsidiarias '" & psAnio & "','" & psMes & "'"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(psSql)
    If Not rs.EOF Then
         ObtieneRepSBS13InversionesCapitalSocialSubsidiarias = rs!nSaldoFinImporte
    Else
         ObtieneRepSBS13InversionesCapitalSocialSubsidiarias = 0
    End If
    dbConec.CierraConexion
Exit Function
ERRORObtiene:
    Err.Raise Err.Number, "ObtieneRepSBS13InversionesCapitalSocialSubsidiarias", Err.Description
End Function
Public Function ObtieneRepSBS13TenenciaOro(ByVal psAnio As String, ByVal psMes As String) As Double
    Dim rs As ADODB.Recordset
On Error GoTo ERRORObtiene
    psSql = "stp_sel_ERS0242015_ObtieneRepSBS13TenenciaOro '" & psAnio & "','" & psMes & "'"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(psSql)
    If Not rs.EOF Then
         ObtieneRepSBS13TenenciaOro = rs!nSaldoFinImporte
    Else
         ObtieneRepSBS13TenenciaOro = 0
    End If
    dbConec.CierraConexion
Exit Function
ERRORObtiene:
    Err.Raise Err.Number, "ObtieneRepSBS13TenenciaOro", Err.Description
End Function
Public Function ObtieneRepSBS13InversionMueblesInmuebles(ByVal psAnio As String, ByVal psMes As String) As Double
    Dim rs As ADODB.Recordset
On Error GoTo ERRORObtiene
    psSql = "stp_sel_ERS0242015_ObtieneRepSBS13InversionMueblesInmuebles '" & psAnio & "','" & psMes & "'"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(psSql)
    If Not rs.EOF Then
         ObtieneRepSBS13InversionMueblesInmuebles = rs!nSaldoFinImporte
    Else
         ObtieneRepSBS13InversionMueblesInmuebles = 0
    End If
    dbConec.CierraConexion
Exit Function
ERRORObtiene:
    Err.Raise Err.Number, "ObtieneRepSBS13InversionMueblesInmuebles", Err.Description
End Function
Public Function ObtieneRepSBS13LimitePosicionGlobal(ByVal psAnio As String, ByVal psMes As String, ByVal psCtaCont As String) As Double
    Dim rs As ADODB.Recordset
On Error GoTo ERRORObtiene
    psSql = "stp_sel_ERS0242015_ObtieneRepSBS13LimitePosicionGlobal '" & psAnio & "','" & psMes & "','" & psCtaCont & "'"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(psSql)
    If Not rs.EOF Then
         ObtieneRepSBS13LimitePosicionGlobal = rs!nSaldoFinImporte
    Else
         ObtieneRepSBS13LimitePosicionGlobal = 0
    End If
    dbConec.CierraConexion
Exit Function
ERRORObtiene:
    Err.Raise Err.Number, "ObtieneRepSBS13LimitePosicionGlobal", Err.Description
End Function
Public Function ObtieneRepSBS13FinanciamientoPersonasVincEmpresa(ByVal pnTpoCambio As Currency, ByVal pdFecha As Date) As Double
    Dim rs As ADODB.Recordset
On Error GoTo ERRORObtiene
    psSql = "stp_sel_ERS0242015_ObtieneRepSBS13FinanciamientoPersonasVincEmpresa " & pnTpoCambio & ",'" & Format(pdFecha, "YYYYMMdd") & "'"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(psSql)
    If Not rs.EOF Then
         ObtieneRepSBS13FinanciamientoPersonasVincEmpresa = rs!nMonto
    Else
         ObtieneRepSBS13FinanciamientoPersonasVincEmpresa = 0
    End If
    dbConec.CierraConexion
Exit Function
ERRORObtiene:
    Err.Raise Err.Number, "ObtieneRepSBS13FinanciamientoPersonasVincEmpresa", Err.Description
End Function
Public Function ObtieneRepSBS13CredDirecTrabEmpresa(ByVal pdFecha As Date) As Double
    Dim rs As ADODB.Recordset
On Error GoTo ERRORObtiene
    psSql = "stp_sel_ERS0242014_ObtieneRepSBS13CredDirecTrabEmpresa '" & Format(pdFecha, "YYYYMMdd") & "'"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(psSql)
    If Not rs.EOF Then
         ObtieneRepSBS13CredDirecTrabEmpresa = rs!nMonto
    Else
         ObtieneRepSBS13CredDirecTrabEmpresa = 0
    End If
    dbConec.CierraConexion
Exit Function
ERRORObtiene:
    Err.Raise Err.Number, "ObtieneRepSBS13CredDirecTrabEmpresa", Err.Description
End Function
Public Function ObtieneRepSBS13FinanciamientoDirectoeIndirectoEmpSistFinan(ByVal pdFecha As Date) As ADODB.Recordset
    Dim rs As ADODB.Recordset
On Error GoTo ERRORObtiene
    psSql = "stp_sel_ERS0242015_ObtieneRepSBS13FinanciamientoDirectoeIndirectoEmpSistFinan '" & Format(DatePart("YYYY", pdFecha), "0000") & "','" & Format(DatePart("M", pdFecha), "00") & "'"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(psSql)
    If Not rs.EOF Then
         Set ObtieneRepSBS13FinanciamientoDirectoeIndirectoEmpSistFinan = rs
    Else
         Set ObtieneRepSBS13FinanciamientoDirectoeIndirectoEmpSistFinan = Nothing
    End If
    dbConec.CierraConexion
Exit Function
ERRORObtiene:
    Err.Raise Err.Number, "ObtieneRepSBS13FinanciamientoDirectoeIndirectoEmpSistFinan", Err.Description
End Function
Public Function ObtieneRepSBS13CredDirecTrabEmpresaMax(ByVal pdFecha As Date) As ADODB.Recordset
    Dim rs As ADODB.Recordset
On Error GoTo ERRORObtiene
    psSql = "stp_sel_ERS0242014_ObtieneRepSBS13CredDirecTrabEmpresaMax '" & Format(pdFecha, "YYYYMMdd") & "'"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(psSql)
    If Not rs.EOF Then
         Set ObtieneRepSBS13CredDirecTrabEmpresaMax = rs
    Else
         Set ObtieneRepSBS13CredDirecTrabEmpresaMax = Nothing
    End If
    dbConec.CierraConexion
Exit Function
ERRORObtiene:
    Err.Raise Err.Number, "ObtieneRepSBS13CredDirecTrabEmpresaMax", Err.Description
End Function
Public Function RegistrarConfigRepSBS13(ByVal psMes As String, ByVal psAnio As String, ByVal psMovNro As String, ByVal pnVig As Integer) As Integer
    On Error GoTo ErrorRegistra
    Dim rs As ADODB.Recordset
    psSql = "stp_ins_ERS0242015_ConfigRepSBS13 '" & psMes & "','" & psAnio & "','" & psMovNro & "'"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(psSql)
        If Not (rs.EOF And rs.BOF) Then
                RegistrarConfigRepSBS13 = rs!nIdRep
        End If
    dbConec.CierraConexion
    Exit Function
ErrorRegistra:
    Err.Raise Err.Number, "Error al Registrar la Configuración", Err.Description
End Function
Public Function RegistrarConfigRepSBS13Det(ByVal pnIdRep As Integer, ByVal psConfRepDesc As String, ByVal pnConfRepVal1 As Double, ByVal pnConfRepVal2 As Double, ByVal pnConfRepResultPorcent As Double)
On Error GoTo ErrorRegistra
    dbConec.AbreConexion
    psSql = "stp_ins_ERS0242015_ConfigRepSBS13Det " & pnIdRep & ",'" & psConfRepDesc & "'," & pnConfRepVal1 & "," & pnConfRepVal2 & "," & pnConfRepResultPorcent
    dbConec.Ejecutar psSql
    Exit Function
ErrorRegistra:
    Err.Raise Err.Number, "Error al Registrar la Configuración", Err.Description
End Function
Public Function ExisteConfigRepSBS13(ByVal psAnio As String, ByVal psMes As String) As Boolean
      Dim rs As ADODB.Recordset
    psSql = "stp_sel_ERS0242015_ExisteConfigRepSBS13Det '" & psMes & "','" & psAnio & "'"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(psSql)
    If Not (rs.EOF And rs.BOF) Then
        ExisteConfigRepSBS13 = IIf(rs!Existe <> 0, True, False)
    Else
        ExisteConfigRepSBS13 = False
    End If
    dbConec.CierraConexion
    Exit Function
End Function
Public Function ObtieneValorxRepSBS13(ByVal psAnio As String, ByVal psMes As String, ByVal psConfDesc As String) As String
    Dim rs As ADODB.Recordset
    psSql = "stp_sel_ERS0872014_ObtieneValorxRepSBS13Det '" & psAnio & "','" & psMes & "','" & psConfDesc & "'"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(psSql)
     If Not (rs.EOF And rs.BOF) Then
        ObtieneValorxRepSBS13 = rs!Valor
    Else
        ObtieneValorxRepSBS13 = ""
    End If
    dbConec.CierraConexion
    Exit Function
End Function
Public Function ObtieneValorxRepSBS13Ratios(ByVal psAnio As String, ByVal psMes As String) As ADODB.Recordset
    dbConec.AbreConexion
    psSql = "stp_sel_ERS0872014_ObtieneValorxRepSBS13DetRatios '" & psAnio & "','" & psMes & "'"
    Set ObtieneValorxRepSBS13Ratios = dbConec.CargaRecordSet(psSql)
     dbConec.CierraConexion
    Exit Function
End Function
Public Function ActualizaConfigRepSBS13(ByVal psMes As String, ByVal psAnio As String, ByVal psMovNro As String, ByVal pnVig As Integer) As Integer
    On Error GoTo ErrorRegistra
    Dim rs As ADODB.Recordset
    psSql = "stp_upd_ERS0242015_ConfigRepSBS13 '" & psMes & "','" & psAnio & "','" & psMovNro & "'"
    dbConec.AbreConexion
    Set rs = dbConec.CargaRecordSet(psSql)
        If Not (rs.EOF And rs.BOF) Then
                ActualizaConfigRepSBS13 = rs!nIdRep
        End If
    dbConec.CierraConexion
    Exit Function
ErrorRegistra:
    Err.Raise Err.Number, "Error al Actualizar la Configuración", Err.Description
End Function
Public Function ActualizaConfigRepSBS13Det(ByVal pnIdRep As Integer, ByVal psConfRepDesc As String, ByVal pnConfRepVal1 As Double, ByVal pnConfRepVal2 As Double, ByVal pnConfRepResultPorcent As Double)
On Error GoTo ErrorRegistra
    dbConec.AbreConexion
    psSql = "stp_upd_ERS0242015_ConfigRepSBS13Det " & pnIdRep & ",'" & psConfRepDesc & "'," & pnConfRepVal1 & "," & pnConfRepVal2 & "," & pnConfRepResultPorcent
    dbConec.Ejecutar psSql
    Exit Function
ErrorRegistra:
    Err.Raise Err.Number, "Error al Registrar la Configuración", Err.Description
End Function
'END PASI
'ALPA 20150901**********************************************************************************
Public Function ObtenerCtaContSaldoBalanceFinal(psCtaContCod As String, pdFecDiaProceso As Date, psMoneda As String, pnTipoCambio As Currency) As Currency
On Error GoTo ObtenerCtaContSaldoBalanceFinalErr
   Dim oRs As ADODB.Recordset
   Dim oConec As DConecta
   Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ObtenerDatosBalanceFinal '" & psCtaContCod & "','" & Format(pdFecDiaProceso, "YYYY/MM/DD") & "','" & psMoneda & "'," & pnTipoCambio
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then
   
      Do While Not oRs.EOF
        nSaldo = oRs!nSaldoFinImporte
           oRs.MoveNext
      Loop
   Else
    nSaldo = 0
   End If
   oConec.CierraConexion
   ObtenerCtaContSaldoBalanceFinal = nSaldo
Exit Function
ObtenerCtaContSaldoBalanceFinalErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerCtaContSaldoBalanceFinal Method")
End Function
'***********************************************************************************************
'PASI 20170727**********************************************************************************
Public Function CargaFTPBSI(psFecha As String, pnNroDigitos As Integer, pbConCaluloUtilidad As Boolean, psCMAC As String) As Recordset
    Dim lsSub As String
    If pbConCaluloUtilidad = False Then
        psSql = "SELECT B.*,C.nSector from BSI B, CtaContBSI C  Where datediff(dd,B.dFecha,'" & psFecha & "') = 0 and "
        psSql = psSql & "LEN(ltrim(rtrim(B.cCtaContCodAnt)))<=22 and (not (B.nSaldoFinImporte=0 and B.nSaldoIniImporte=0) or SubString(B.cCtaContCod,10,4) like '1924') "
        psSql = psSql & "   and ((B.cCtaContCodAnt=left(C.cCtaContBSI,2) and len(C.cCtaContBSI)<=3) or ((B.cCtaContCodAnt=replace(C.cCtaContBSI,'M',B.cMoneda) and len(C.cCtaContBSI)>2))) "
        psSql = psSql & "   Order By SubString(B.cCtaContCod,12,1),SubString(B.cCtaContCod,1,11) + SubString(B.cCtaContCod,13,17) + B.cMoneda "
    Else
        psSql = "exec stp_sel_BSI_GeneraBSIConUtilidad '" & Format(psFecha, "YYYY/MM/DD") & "', " & pnNroDigitos
    End If
    Set CargaFTPBSI = dbConec.CargaRecordSet(psSql)
End Function
Public Sub InsertaFTPBSI(pdFecha As Date, psCodigo, pnCateBala As Integer, pnMoneda As Integer)
On Error GoTo InsertaFTPBSIErr
   
  psSql = "INSERT BSI " _
        & "SELECT '" & Format(pdFecha, gsFormatoFecha) & "', SUBSTRING('" & psCodigo & "' + c.cCtaContCod + '" & String(20, "0") & "',1,29), SUM(z.nSaldoIniImporte ) as nSaldoIni, " _
        & "       SUM( z.nDebe ) as nDebe, SUM( z.nHaber) as nHaber, " _
        & "       ISNULL(SUM(CASE WHEN cCtaClasifica = 'D' THEN z.nSaldoIniImporte + z.nDebe - z.nHaber " _
        & "                       WHEN cCtaClasifica = 'A' THEN z.nSaldoIniImporte - z.nDebe + z.nHaber END ) ,0) as nSaldoFinImporte, cBalanceTipo, c.cCtaContcod " _
        & "FROM   BalanceTemp as z JOIN " & sCentralCom & "CtaCont c ON z.cCtaContCod LIKE c.cCtaContCod + '%' " _
        & "WHERE  cBalanceCate = '" & pnCateBala & "' and cBalanceTipo = '" & pnMoneda & "' and c.cCtaContCod not like '[678]%'" _
        & "GROUP BY c.cCtaContCod, c.cCtaContDesc, cBalanceTipo " _
        & "ORDER BY c.cCtaContCod "
  dbConec.CommadTimeOut = 7200
  dbConec.Ejecutar psSql
Exit Sub
InsertaFTPBSIErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:InsertaBSI Method")
End Sub
Public Sub ActualizaFTPBSI(pdFecha As Date, psCodigo, pnCateBala As Integer, pnMoneda As Integer)
On Error GoTo ActualizaFTPBSIErr
   psSql = "UPDATE BSI Set cCtaContCod = '" & psCodigo & "' + cCtaContCodAnt + '0" & pnMoneda & "00000000000000000' " _
       & " WHERE datediff(dd,dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and " _
       & "       LEN(ltrim(rtrim(cCtacontcodant))) = 1 and cMoneda = '" & pnMoneda & "'"
   dbConec.Ejecutar psSql
                                              
   psSql = "UPDATE BSI Set cCtaContCod = '" & psCodigo & "' + cCtaContCodAnt + '" & pnMoneda & "00000000000000000' " _
       & " WHERE datediff(dd,dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and " _
       & " LEN(ltrim(rtrim(cCtacontcodant))) = 2 and cMoneda = '" & pnMoneda & "'"
   dbConec.Ejecutar psSql
Exit Sub
ActualizaFTPBSIErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ActualizaFTPBSI Method")
End Sub
Public Sub EliminaBSI(pdFecha As Date, Optional psCodigo As String = "")
On Error GoTo EliminaBSIErr
   psSql = "DELETE BSI " _
         & "WHERE dFecha = '" & Format(pdFecha, gsFormatoFecha) & "' " & IIf(psCodigo = "", "", " and cCtaContCod = '" & psCodigo & "'")
   dbConec.CommadTimeOut = 7200
   dbConec.Ejecutar psSql
Exit Sub
EliminaBSIErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:EliminaBSI Method")
End Sub
Public Function CargarBSI_Sector(psDato As String) As Recordset
   Dim lsSub As String
   psSql = "exec stp_sel_BSI_Sector_Periodo '" & psDato & "'"
   Set CargarBSI_Sector = dbConec.CargaRecordSet(psSql)
End Function
Public Function CargarSectorCreditos(ByVal pdFecha As Date) As Integer
    psSql = "stp_sel_SectorCreditos '" & Format(pdFecha, "YYYYMMdd") & "'"
    dbConec.CargaRecordSet (psSql)
    CargarSectorCreditos = 0
End Function
'PASI END***************************************************************************************

Public Function ObtieneReporteIntDifAmp_NoAmp(ByVal pdFecha As Date, psTipo As String) As ADODB.Recordset
On Error GoTo OtieneReporteIntDifAmp_NoAmpErr
   Dim oRs As New ADODB.Recordset
   Dim oConec As New DConecta
   oConec.AbreConexion
   psSql = "exec stp_sel_ObtieneReporteGeneralIntDiferidoAmp_NoAmp '" & Format(pdFecha, "yyyymmdd") & "', '" & psTipo & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtieneReporteIntDifAmp_NoAmp = oRs
   oConec.CierraConexion
Exit Function
OtieneReporteIntDifAmp_NoAmpErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:OtieneReporteIntDifAmp_NoAmp Method")
End Function 'NAGL 202008 Según Acta N°063-2020

Public Function ObtieneCuentasBalanceANX16(ByVal cMoneda As String) As ADODB.Recordset
On Error GoTo ObtieneCuentasBalanceANX16Err
   Dim oRs As New ADODB.Recordset
   Dim oConec As New DConecta
   oConec.AbreConexion
   psSql = "exec stp_ins_cargarCuentasBalanceANX16 '" & cMoneda & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtieneCuentasBalanceANX16 = oRs
   oConec.CierraConexion
Exit Function
ObtieneCuentasBalanceANX16Err:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ObtieneCuentasBalanceANX16 Method")
End Function 'NAGL 202012 Según ACTA N°094-2020

Public Function EvaluaCuentaLikeAnx16(ByVal cCtaCnt As String) As Boolean
On Error GoTo EvaluaCuentaLikeAnx16Err
   Dim oRs As New ADODB.Recordset
   Dim oConec As New DConecta
   Dim pbValor As Boolean
   oConec.AbreConexion
   psSql = "exec stp_sel_EvaluaCuentaLike_ANX16 '" & cCtaCnt & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
    EvaluaCuentaLikeAnx16 = oRs!pbValor
   oConec.CierraConexion
Exit Function
EvaluaCuentaLikeAnx16Err:
   Call RaiseError(MyUnhandledError, "DBalanceCont:EvaluaCuentaLikeAnx16 Method")
End Function 'NAGL 202012 Según ACTA N°094-2020

Public Function ObtenerCtasxPagarOpeReactivaxTramos(ByVal pdFecDiaProceso As Date, ByVal psMoneda As String) As ADODB.Recordset
    On Error GoTo ObtenerCtasxPagarOpeReactivaxTramosErr
    Dim oRs As ADODB.Recordset
    Dim oConec As DConecta
    Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
    psSql = "exec stp_sel_CtasxPagarOpeReactiva_Tramos '" & Format(pdFecDiaProceso, "yyyymmdd") & "','" & psMoneda & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   Set ObtenerCtasxPagarOpeReactivaxTramos = oRs
   oConec.CierraConexion
    Exit Function
ObtenerCtasxPagarOpeReactivaxTramosErr:
    Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerCtasxPagarOpeReactivaxTramos Method")
End Function 'NAGL 202012 Según ACTA N°094-2020

'JIPR20210318 REPOGARTR INICIO
Public Function SaldosRepoGartAnexo(ByVal psTipo As String, ByVal psMoneda As String, ByVal psRepoGart As String) As Currency
    On Error GoTo SaldosRepoGartAnexoErr
    Dim oRs As ADODB.Recordset
    Dim oConec As DConecta
    Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
    psSql = "exec stp_sel_SaldosRepoGartAnexo '" & psTipo & "','" & psMoneda & "','" & psRepoGart & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then

        Do While Not oRs.EOF
            nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
            oRs.MoveNext
        Loop
    Else
        nSaldo = 0
    End If
    oConec.CierraConexion
    SaldosRepoGartAnexo = nSaldo
    Exit Function
SaldosRepoGartAnexoErr:
    Call RaiseError(MyUnhandledError, "DBalanceCont:SaldosRepoGartAnexo Method")
End Function
'JIPR20210318 REPOGARTR FIN

'JIPR20211005  INICIO
Public Function ObtenerTotalEncajeAcumuladoFechAnx15A(ByVal pdFechaGenera As Date) As Currency
    On Error GoTo ObtenerTotalEncajeAcumuladoFechAnx15AErr
    Dim oRs As ADODB.Recordset
    Dim oConec As DConecta
    Dim nSaldo As Currency
   Set oRs = New ADODB.Recordset
   Set oConec = New DConecta
   oConec.AbreConexion
    psSql = "exec sp_sel_ObtenerTotalEncajeAcumuladoFechAnx15A  '" & Format(pdFechaGenera, "yyyymmdd") & "'"
   Set oRs = oConec.CargaRecordSet(psSql)
   If Not oRs.BOF And Not oRs.EOF Then

        Do While Not oRs.EOF
            nSaldo = IIf(IsNull(oRs!nSaldo), 0, oRs!nSaldo)
            oRs.MoveNext
        Loop
    Else
        nSaldo = 0
    End If
    oConec.CierraConexion
    ObtenerTotalEncajeAcumuladoFechAnx15A = nSaldo
    Exit Function
ObtenerTotalEncajeAcumuladoFechAnx15AErr:
    Call RaiseError(MyUnhandledError, "DBalanceCont:ObtenerTotalEncajeAcumuladoFechAnx15A Method")
End Function
'JIPR20211005  FIN


