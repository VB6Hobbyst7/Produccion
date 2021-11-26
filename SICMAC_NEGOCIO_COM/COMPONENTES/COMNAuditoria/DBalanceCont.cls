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
Dim dbConec As COMConecta.DCOMConecta
Dim psSql As String
Dim prs As ADODB.Recordset
Dim sConexion As String
Dim sCentralCom As String

Private Sub Class_Initialize()
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Set oImp = New COMDConstSistema.DCOMImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
    'Dim oIni As New COMConecta.DCOMClasIni
    'sCentralCom = oIni.BaseComunes
    'Set oIni = Nothing
    Set dbConec = New COMConecta.DCOMConecta
    If Not dbConec.AbreConexion() Then
        MsgBox Err.Description, vbCritical
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
   MsgBox Err.Description, vbCritical
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
   MsgBox Err.Description, vbCritical
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
   MsgBox Err.Description, vbCritical
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
         sCondBala = "and ( (Substring(mc.cCtaContCod,3,1) Like '[13]' or substring(mc.cCtaContCod,1,1) IN (" & sCtaMN & ") " _
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
        & "WHERE M.nMovEstado = '" & gMovEstContabMovContable & "' and not M.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & "," & gMovFlagExtornado & ") and SubString(m.cMovNro,1,8) BETWEEN '" & psFechaDel & "' and '" & psFechaAl & "' " _
        & "      " & sCondBala & " and mc.cCtaContCod IS NOT NULL and mc.cCtaContCod <> '' " & IIf(pbBCient Or pbConCierreAnual, "", " and not M.cOpeCod LIKE '" & Left(gContCierreAnual, 5) & "%' ") _
        & "GROUP BY " & sCta & ", cls.cCtaCaracter "
   dbConec.Ejecutar psSql

Exit Sub
InsertaMovimientosMesErr:
  MsgBox Err.Description, vbCritical
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
   'dbConec.CommadTimeOut = 7200
   dbConec.Ejecutar psSql
Exit Sub
InsertaMovimientosMesErr:
   MsgBox Err.Description, vbCritical
End Sub

Public Function EjecutaBatch() As Integer
On Error GoTo EjecutaBatchErr
    EjecutaBatch = dbConec.EjecutarBatch
    Exit Function
EjecutaBatchErr:
    MsgBox Err.Description, vbCritical
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
   MsgBox Err.Description, vbCritical
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
   MsgBox Err.Description, vbCritical
End Sub
