VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAjusteCont"
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

Dim pnBitCentral As Boolean
Dim sServidorConsolidada As String
Dim sCodEmp As String


Public Function CargaAjusteInflacion(Optional psCtaCod As String = "", Optional psAjusteCod As String = "", Optional psFiltro As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
Dim sCond As String
   On Error GoTo CargaAjusteInflacionErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion Then
      If psCtaCod <> "" Then
         sCond = "WHERE cCtaContCod ='" & psCtaCod & "' "
      End If
      If psAjusteCod <> "" Then
         sCond = sCond & IIf(sCond = "", "WHERE", "AND") & " cAjusteCod = '" & psAjusteCod & "' "
      End If
      If psFiltro <> "" Then
         sCond = sCond & IIf(sCond = "", "WHERE", "AND") & " " & psFiltro
      End If
      psSql = "SELECT cCtaContCod, cAjusteCod, cAjusteDescrip, Convert(varchar(10),dAjusteFecha,103) dAjusteFecha, nAjusteValor1, nAjusteValor2, nAjusteValor3 From AjusteInflacion " & sCond & " Order By cCtaContCod, cAjusteCod "
      Set CargaAjusteInflacion = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaAjusteInflacionErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:CargaAjusteInflacion Method")
End Function

Public Sub InsertaAjuste(psCtaCod As String, psAjusteCod As String, psDescrip As String, psFecha As String, pnValor1 As Currency, pnValor2 As Currency, pnValor3 As Currency)
   On Error GoTo InsertaAjusteErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion Then
      psSql = "INSERT AjusteInflacion (cCtaContCod, cAjusteCod, cAjusteDescrip, dAjusteFecha, nAjusteValor1, nAjusteValor2, nAjusteValor3) " _
            & "VALUES ('" & psCtaCod & "','" & psAjusteCod & "','" & psDescrip & "', '" & psFecha & "'," & pnValor1 & "," & pnValor2 & "," & pnValor3 & ")"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaAjusteErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:InsertaAjuste Method")
End Sub

Public Sub ActualizaAjuste(psCtaCod As String, psAjusteCod As String, psDescrip As String, psFecha As String, pnValor1 As Currency, pnValor2 As Currency, pnValor3 As Currency)
   On Error GoTo ActualizaAjusteErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion Then
      psSql = "UPDATE AjusteInflacion SET cAjusteDescrip = '" & psDescrip & "', dAjusteFecha = '" & psFecha & "', nAjusteValor1  = " & pnValor1 & ", nAjusteValor2  = " & pnValor2 & ", nAjusteValor3  = " & pnValor3 & " WHERE cCtaContCod = '" & psCtaCod & "' and cAjusteCod = '" & psAjusteCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizaAjusteErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:ActualizaAjuste Method")
End Sub

Public Sub EliminaAjuste(psCtaCod As String, psAjusteCod As String)
   On Error GoTo EliminaAjusteErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion Then
      psSql = "DELETE AjusteInflacion WHERE cCtaContCod = '" & psCtaCod & "' and cAjusteCod = '" & psAjusteCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaAjusteErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:EliminaAjuste Method")
End Sub

Public Sub TrasladoValorAjustado()
Dim ldFecha As Date
   On Error GoTo TrasladoValorAjustadoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion Then
      psSql = "SELECT Min(dAjusteFecha) dAjusteFecha FROM AjusteInflacion "
      Set prs = dbConec.CargaRecordSet(psSql)
      If Not prs.EOF Then
         ldFecha = CDate(Left(Format(prs!dAjusteFecha, gsFormatoFechaView), 6) & Format(Year(prs!dAjusteFecha) + 1, "0000"))
         psSql = "update aj set nAjusteValor1 = aj.nAjusteValor2, nAjusteValor2 = aj.nAjusteValor3, nAjusteValor3 = round( nAjusteValor3 * round(dic.nValor/ ipm.nValor,3), 2), dAjusteFecha = '" & Format(ldFecha, gsFormatoFecha) & "' " _
               & "from AjusteInflacion aj join ipm on year(ipm.dFecha) = year(aj.dAjusteFecha) and month(ipm.dFecha) = month(aj.dAjusteFecha), " _
               & "     ipm dic " _
               & "where dic.dFecha = '" & Format(ldFecha, gsFormatoFecha) & "' "
         dbConec.Ejecutar psSql
      End If
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
TrasladoValorAjustadoErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:TrasladoValorAjustado Method")
End Sub

Public Function CargaIPM(Optional psFecha As String = "", Optional psFiltro As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
Dim sCond As String
   On Error GoTo CargaImpuestoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion Then
      If psFecha <> "" Then
         sCond = "WHERE datediff(m,dFecha,'" & psFecha & "') = 0 "
      End If
      If psFiltro <> "" Then
         sCond = sCond & IIf(psFecha = "", "WHERE", "AND") & " " & psFiltro
      End If
      psSql = "SELECT Convert(varchar(10),dFecha,103) dFecha, Convert(decimal(16,6),nValor) nValor From IPM " & sCond & " Order By convert(datetime,dFecha) "
      Set CargaIPM = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaImpuestoErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:CargaImpuesto Method")
End Function

Public Sub InsertaIPM(psFecha As String, pnValor As Double)
   On Error GoTo InsertaIPMErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion Then
      psSql = "INSERT IPM (dFecha, nValor) VALUES ('" & psFecha & "'," & pnValor & ")"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaIPMErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:InsertaIPM Method")
End Sub

Public Sub ActualizaIPM(psFecha As String, pnValor As Double)
   On Error GoTo ActualizaIPMErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "UPDATE IPM SET nValor  = " & pnValor & " WHERE dFecha = '" & psFecha & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizaIPMErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:ActualizaIPM Method")
End Sub

Public Sub EliminaIPM(psFecha As String)
   On Error GoTo EliminaIPMErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE IPM WHERE dFecha = '" & psFecha & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaIPMErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:EliminaIPM Method")
End Sub

Public Function CargaAjusteHistorico() As Recordset
Dim sCond As String
   On Error GoTo CargaAjusteHistoricoErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then

      psSql = "SELECT da.*, c.cCtaContDesc, cls.cCtaCaracter as cTipoDH " _
           & "FROM AjusteInflacion da JOIN " & sCentralCom & "CtaCont c ON c.cCtaContCod = da.cCtaContCod " _
           & "     JOIN " & sCentralCom & "CtaContClase CLS on da.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' order by da.cCtaContCod "
      Set CargaAjusteHistorico = dbConec.CargaRecordSet(psSql)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaAjusteHistoricoErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:CargaAjusteHistorico Method")
End Function

Public Function VerAjusteCierreAnual(psFechaCie As String, psCond As String) As Recordset
On Error GoTo VerAjusteCierreAnualErr
   Set dbConec = New DConecta
   dbConec.AbreConexion
   If pnBitCentral = True Then
        psSql = "SELECT c.cCtaContCod, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte, cls.cCtaCaracter cTipo " _
            & "FROM " & sCentralCom & "CtaCont c LEFT JOIN " _
            & "     (SELECT cs.cCtaContCod, cs.nCtaSaldoImporte FROM CtaSaldo cs WHERE dCtaSaldoFecha = " _
            & "      ( SELECT MAX(a.dCtaSaldoFecha) FROM CtaSaldo a " _
            & "        WHERE  a.cCtaContCod = cs.cCtaContCod and dCtasaldoFecha <= '" & psFechaCie & "') " _
            & "     ) cs on cs.cCtaContCod = c.cCtaContCod, " & sCentralCom & "CtaContClase cls " _
            & "WHERE " & psCond _
            & "     and c.cCtaContCod LIKE cls.cCtaContCod + '%'  "
    Else
        psSql = "SELECT c.cCtaContCod, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte, cls.cCtaCaracter cTipo " _
             & "FROM CtaCont c LEFT JOIN CtaContClase cls ON c.cCtaContCod LIKE cls.cCtaContCod + '%' " _
             & "     LEFT JOIN CtaSaldo cs ON cs.cCtaContCod = c.cCtaContCod " _
             & "         and cs.dCtaSaldoFecha = ( SELECT MAX(a.dCtaSaldoFecha) FROM CtaSaldo a " _
             & "        WHERE  a.cCtaContCod = cs.cCtaContCod and a.dCtasaldoFecha <= '" & psFechaCie & "') " _
             & "WHERE " & psCond & " ORDER BY c.cCtaContCod  " _
             & " " _
             & "  "
    End If
    
   Set VerAjusteCierreAnual = dbConec.CargaRecordSet(psSql)
   Set dbConec = Nothing
Exit Function
VerAjusteCierreAnualErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:VerAjusteCierreAnual Method")
End Function

Public Function GetCuentasAjusteBalance(psCtaCod As String, pnAnio As Integer, pnMes As Integer, psOpecod As String) As ADODB.Recordset
On Error GoTo GetCuentasAjusteBalanceErr
   Set dbConec = New DConecta
   dbConec.AbreConexion
psSql = "SELECT c.cctacontcod, c.cctacontdesc " _
     & "FROM CtaCont c JOIN ( SELECT DISTINCT LEFT(cCtaContCod,2)+'6'+SubString(cCtaContCod,4,22) cCtaContCod FROM BalanceEstad WHERE cBalanceAnio = '" & pnAnio & "' and cBalanceMes = '" & Format(pnMes, "00") & "' and cBalanceCate = '1' and nSaldoFinImporte <> 0 " _
     & "     ) Bal ON Bal.cCtaContCod = c.cCtaContCod, " _
     & "CtaCont c1 " _
     & "WHERE c.cctacontcod LIKE '" & psCtaCod & "6%' and c1.cctacontcod like c.cctacontcod + '%' " _
     & "      and not c.cCtaContCod in (SELECT cCtaContCod FROM OpeCta WHERE cOpeCod = '" & psOpecod & "' and cOpeCtaOrden = '1' ) " _
     & "GROUP BY c.cCtaContCod , c.cCtaContDesc " _
     & "HAVING count(*) = 1 "
Set GetCuentasAjusteBalance = dbConec.CargaRecordSet(psSql)
   Set dbConec = Nothing
Exit Function
GetCuentasAjusteBalanceErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'****************************
'** NUEVAS FUNCIONES
'****************************

Public Function AjusteReclasificaCartera(psFecha As String, pnMoneda As Integer, psClase As String, pnTipCambio As Currency) As Recordset
On Error GoTo AjusteReclasificaCarteraErr
   Set dbConec = New DConecta
   dbConec.AbreConexion
'Dim lscadser As String
'Dim lscodemp As String
'Dim oConst As NConstSistemas
'Set oConst = New NConstSistemas
'lscodemp = oConst.LeeConstSistema(gConstSistCodCMAC)
'lscadser = oConst.LeeConstSistema(gConstSistServCentralRiesgos)
'Set oConst = Nothing

If pnBitCentral = True Then
    If sCodEmp = "112" Or sCodEmp = "108" Or sCodEmp = "109" Then 'TRUJILLO /ICA /HUANCAYO
    ' ** Final
'        psSql = "SELECT cred.cCtaContcod cta1, ISNULL(cred.nSaldo,0) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte, IsNull(cred.cCtaContcod,cs.cCtaContcod) cCtaContCodOrd  " _
'             & "FROM (  SELECT cCtaContCod, SUM(nSaldo) nSaldo FROM " _
'             & " ( Select '14' + cMoneda + " _
'             & "  CASE WHEN cRefinanc = 'R' and cEstado IN ('1','3') THEN '4' ELSE cEstado END " _
'             & "  + '0' + left(cProdCod,1) " _
'             & "  + CASE WHEN cRFA = 'RFC' AND CESTADO IN ('1','3') THEN '25010106' WHEN cRFA ='DIF' AND CESTADO IN ('1','3') THEN '25010206' " _
'             & "         WHEN cRFA = 'RFC' AND CESTADO IN ('5','6') THEN '19020106' WHEN cRFA ='DIF' AND CESTADO IN ('5','6') THEN '19020206' " _
'             & "         WHEN cProdCod = '320' THEN '2009' WHEN cProdCod = '423' THEN '23' WHEN cProdCod = '305' THEN '13' " _
'             & "         WHEN cRefinanc = 'R' and cEstado IN ('5') THEN '19' " _
'             & "              + CASE WHEN LEFT(cProdCod,1) IN ('1','2') THEN '' else '' END + '06' " _
'             & "     Else '06' END + CASE WHEN cRefinanc = 'R' and cEstado IN ('5') THEN '' Else Case When cProdCod In ('305','320') Then '' When left(cProdCod,1) = '3' Then '09' Else Case When cProdCod = '401' Then '01' else '02' End End End " _
'             & "  + Case when cProdCod In ('320','305') then '' else '0'+ cPlazo +  (SELECT LTRIM(RTRIM(cEquivAsiento)) From AsientoProdEquiv AE Where cEquivTpo ='SC' and AE.cProducto = cProdCod) " _
'             & "  +  (Select cCtaCont From dbconsolidada..coloclineacreditoequiv where clineacred = cFuenteFinanc) " _
'             & "  end + cAgeCod as cCtaContCod, " _
'             & "  CASE WHEN cMoneda = '" & gMonedaExtranjera & "' THEN ROUND(nSaldo * " & pnTipCambio & ",2) ELSE nSaldo END nSaldo " _
'             & "  FROM " & sServidorConsolidada & "CredSaldosCont WHERE dFecha = '" & psFecha & "' and cConcepto = 'C' and cMoneda = '" & pnMoneda & "' " _
'             & "  ) Est GROUP BY cCtaContCod " _
'             & " ) Cred FULL OUTER JOIN " _
'             & "( SELECT cCtaContCod, nCtaSaldoImporte " _
'             & "  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psClase & pnMoneda & "[1234567]%' and cs.dCtaSaldoFecha = " _
'             & "  CONVERT(DATETIME, (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & Format(psFecha, "MM/dd/YYYY") & "' ),102) " _
'             & ") cs on cs.cCtaContCod = Cred.cCtaContCod " _
'             & "WHERE ISNULL(nSaldo,0) <> 0 or ISNULL(nCtaSaldoImporte,0) <> 0 ORDER BY cCtaContCodOrd "
    psSql = "exec B2_stp_sel_AjusteReclasificacionCartera '14','" & pnMoneda & "','" & Format(psFecha, "MM/dd/YYYY") & "'," & pnTipCambio
    End If
    'If sCodEmp = "102" Then 'LIMA
    '    psSql = "SELECT cred.cCtaContcod cta1, ISNULL(cred.nSaldo,0) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
             & "FROM ( SELECT cCtaContCod, SUM(nSaldo) nSaldo FROM ( Select '14'+cMoneda+ " _
             & "     CASE WHEN cRefinanc = 'R' and cEstado IN ('1','3') THEN '4' ELSE cEstado END + " _
             & "     '0'+left(cProdCod,1) + CASE WHEN cRefinanc = 'R' and cEstado IN ('5','6') THEN '1929' Else '' END + " _
             & "     ISNULL(CASE WHEN cProdCod = '" & Producto.gColConsuPrendario & "' THEN '13' ELSE (SELECT cEquivAsiento FROM AsientoProdEquiv eq WHERE eq.cProducto = cProdCod and cEquivTpo = 'CD') END,'') + " _
             & "     CASE WHEN cEstado IN ('5') and not cDemandado = 'S' THEN '90' ELSE '' END + " _
             & "     RTRIM(ISNULL((SELECT ISNULL(cEquivAsiento,'') FROM AsientoProdEquiv eq WHERE eq.cProducto = cProdCod and cEquivTpo = 'SC'),''))  cCtaContCod, " _
             & "     CASE WHEN cMoneda = '" & gMonedaExtranjera & "' THEN ROUND(nSaldo * " & pnTipCambio & ",2) ELSE nSaldo END nSaldo " _
             & "   FROM " & sServidorConsolidada & "CredSaldosCont WHERE dFecha = '" & Format(psFecha, "mm/dd/YYYY") & "' and cConcepto = 'C' and cMoneda = '" & pnMoneda & "' " _
             & "   ) Est GROUP BY cCtaContCod " _
             & " ) Cred FULL OUTER JOIN " _
             & "( SELECT cCtaContCod, nCtaSaldoImporte " _
             & "  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psClase & pnMoneda & "[1234567]%' and cs.dCtaSaldoFecha = " _
             & "  CONVERT(DATETIME, (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & Format(psFecha, "MM/dd/YYYY") & "' ),102) " _
             & ") cs on cs.cCtaContCod = Cred.cCtaContCod " _
             & "WHERE ISNULL(nSaldo,0) <> 0 or ISNULL(nCtaSaldoImporte,0) <> 0 ORDER BY cred.cCtaContCod "
    'End If
ElseIf pnBitCentral = False Then
'    If sCodEmp = "112" Then 'TRUJILLO
'        psSql = "SELECT cred.cCtaContcod cta1, ISNULL(cred.nSaldo,0) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
'             & "FROM ( select '14'+cMoneda+ " _
'             & "  CASE WHEN cRefinanc = 'R' and cEstado IN ('1','3') THEN '4' ELSE cEstado END + " _
'             & "  '0'+left(cProdCod,1) + CASE WHEN cProdCod = '305' THEN '13' WHEN cProdCod = '320' THEN '20' WHEN cProdCod = '423' THEN '23' WHEN cRefinanc = 'R' and cEstado IN ('5','6') THEN '19' + CASE WHEN LEFT(cProdCod,1) IN ('1','2') THEN '29' ELSE '' END Else '06' END + " _
'             & "  CASE WHEN cEstado IN ('5') and cDemandado='S' and LEFT(cProdCod,1) <> '3' and cRefinanc = 'R' THEN '29' WHEN cEstado IN ('5') and cDemandado = 'S' and LEFT(cProdCod,1) =  '3' THEN '9'+cPlazo WHEN cEstado IN ('5') and cDemandado = 'S' and not cRefinanc = 'R' THEN '9'+cPlazo ELSE '0'+ cPlazo END + " _
'             & "  CASE WHEN cProdCod IN ('305','320','423') THEN '01' ELSE RIGHT(cProdCod,2) END + " _
'             & "  CASE WHEN cFuenteFinanc = 'A' THEN '10' WHEN cFuenteFinanc = 'B' THEN '11' WHEN cFuenteFinanc = 'C' THEN '12' ELSE '0'+cFuenteFinanc END + " _
'             & "  CASE WHEN cAgeCod in ('99') THEN '11' ELSE cAgeCod END cCtaContCod, " _
'             & "  CASE WHEN cMoneda = '2' THEN ROUND(nSaldo * " & pnTipCambio & ",2) ELSE nSaldo END nSaldo " _
'             & "  FROM " & sServidorConsolidada & "CredSaldosCont WHERE dFecha = '" & psFecha & "' and cConcepto = 'C' and cMoneda = '" & pnMoneda & "' ) Cred FULL OUTER JOIN " _
'             & "( SELECT cCtaContCod, nCtaSaldoImporte " _
'             & "  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psClase & pnMoneda & "[1234567]%' and cs.dCtaSaldoFecha = " _
'             & "  (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' ) " _
'             & ") cs on cs.cCtaContCod = Cred.cCtaContCod " _
'             & "WHERE ISNULL(nSaldo,0) <> 0 or ISNULL(nCtaSaldoImporte,0) <> 0 ORDER BY cred.cCtaContCod "
'    End If
End If

   Set AjusteReclasificaCartera = dbConec.CargaRecordSet(psSql)
   Set dbConec = Nothing
Exit Function
AjusteReclasificaCarteraErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:AjusteReclasificaCartera Method")
End Function


Public Function AjusteInteresesColocaciones(pnTipoAsiento As Integer, psFecha As String, pnMoneda As Integer, psClase As String, pnTipCambio As Currency, psCtaDebe As String, Optional pbAsiRevDev As Boolean = False, Optional pnTipoCambioVenta As Currency = 0, Optional pnTipoCambioCompra As Currency = 0, Optional psRep As String = "") As Recordset
'NAGL 202102 Agregó psRep
On Error GoTo AjusteInteresesColocacionesErr
   'ALPA 20091006********************************************************
   Dim nSaldoTotalProvProc As Currency
   nSaldoTotalProvProc = VerificarSumaProvisionProciclica
   '*********************************************************************
   Set dbConec = New DConecta
   dbConec.AbreConexion

    'Dim lsCadSer As String
    'Dim lscodemp As String
    'Dim oConst As NConstSistemas
    'Set oConst = New NConstSistemas
    'lscodemp = oConst.LeeConstSistema(gConstSistCodCMAC)
    'lsCadSer = oConst.LeeConstSistema(gConstSistServCentralRiesgos)
    'Set oConst = Nothing
 
    If pnTipoAsiento = 1 Then 'Devengados
        If pnBitCentral = True Then
            '** Final
            If sCodEmp = "112" Or sCodEmp = "108" Or sCodEmp = "109" Then 'Trujillo / Ica
'                  psSql = "SELECT cred.cCtaContCod cta1, SUM(ISNULL(cred.nSaldo,0)) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
'                     & "FROM ( select '" & psClase & "'" _
'                     & " + '0'+left(cProdCod,1) " _
'                     & " + CASE WHEN cProdCod = '" & Producto.gColConsuPrestAdm & "' THEN '2009' " _
'                     & "        WHEN cProdCod = '" & Producto.gColHipoMiVivienda & "' THEN '23' WHEN cProdCod = '" & Producto.gColConsuPrendario & "' THEN '13' " _
'                     & "        Else '06' END " _
'                     & " + Case WHEN cProdCod In ('" & Producto.gColConsuPrendario & "','" & Producto.gColConsuPrestAdm & "') then '' else '0'+ cPlazo  " _
'                     & " + CASE WHEN cProdCod ='" & Producto.gColConsuPrendario & "' THEN '01' " _
'                     & "        ELSE (SELECT LTRIM(RTRIM(cEquivAsiento)) From AsientoProdEquiv AE Where cEquivTpo ='SC' and AE.cProducto = cProdCod) End  + (Select cCtaCont From dbconsolidada..coloclineacreditoequiv where clineacred = cFuenteFinanc)  " _
'                     & "  end + cAgeCod as cCtaContCod, " _
'                     & "  nSaldo " _
'                     & "  FROM " & sServidorConsolidada & "CredSaldosCont WHERE cEstado = '1' and cRefinanc = 'N' and dFecha = '" & psFecha & "' and cConcepto = 'I' and cMoneda = '" & pnMoneda & "' ) Cred FULL OUTER JOIN " _
'                     & "( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) nCtaSaldoImporte " _
'                     & "  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
'                     & "  (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' ) " _
'                     & ") cs on cs.cCtaContCod = Cred.cCtaContCod " _
'                     & "WHERE ISNULL(nSaldo,0) <> ISNULL(nCtaSaldoImporte,0) GROUP BY cred.cCtaContCod, cs.cCtaContCod, cs.nCtaSaldoImporte ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod) "
             'ALPA 20120402***********************************************************
             psSql = "exec B2_stp_sel_AjusteInteresesDevengados '14', '" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "'," & pnTipCambio
             'psSql = "exec B2_stp_sel_AjusteInteresesDevengados '14', '" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "'," & pnTipCambio & "," & pnTipoCambioVenta & "," & pnTipoCambioCompra
             '************************************************************************
             End If
               
             'If sCodEmp = "102" Then  'Lima
                  '** Final
             '     psSql = "SELECT cred.cCtaContCod cta1, SUM(ISNULL(cred.nSaldo,0)) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
                     & " FROM ( select '" & psClase & "'+" _
                     & "  '0'+left(cProdCod,1) + ISNULL(CASE WHEN cProdCod = '" & Producto.gColConsuPrendario & "' THEN '13' ELSE (SELECT cEquivAsiento FROM AsientoProdEquiv eq WHERE eq.cProducto = cProdCod and cEquivTpo = 'CD') END,'') + " _
                     & "  RTRIM(ISNULL((SELECT cEquivAsiento FROM AsientoProdEquiv eq WHERE eq.cProducto = cProdCod and cEquivTpo = 'SC'),''))  cCtaContCod, " _
                     & "  nSaldo " _
                     & "  FROM " & sServidorConsolidada & "CredSaldosCont WHERE cEstado = '1' and cRefinanc = 'N' and dFecha = '" & psFecha & "' and cConcepto = 'I' and cMoneda = '" & pnMoneda & "' ) Cred FULL OUTER JOIN " _
                     & "( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) nCtaSaldoImporte " _
                     & "  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
                     & "  (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' ) " _
                     & ") cs on cs.cCtaContCod = Cred.cCtaContCod " _
                     & "WHERE ISNULL(nSaldo,0) <> ISNULL(nCtaSaldoImporte,0) GROUP BY cred.cCtaContCod, cs.cCtaContCod, cs.nCtaSaldoImporte ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod) "
              'End If
        ElseIf pnBitCentral = False Then
            If sCodEmp = "112" Then  ''Trujillo
              'psSql = "SELECT cred.cCtaContCod cta1, SUM(ISNULL(cred.nSaldo,0)) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte FROM ( select '" & psClase & "'+" _
                 & "  '0'+left(cProdCod,1) + CASE WHEN cProdCod = '305' THEN '13' WHEN cProdCod = '320' THEN '20' WHEN cProdCod = '423' THEN '23' Else '06' END + " _
                 & "  CASE WHEN cEstado IN ('5') and cDemandado='S' and LEFT(cProdCod,1) <> '3' and cRefinanc = 'R' THEN '29' WHEN cEstado IN ('5') and cDemandado = 'S' and LEFT(cProdCod,1) =  '3' THEN '9'+cPlazo WHEN cEstado IN ('5') and cDemandado = 'S' and not cRefinanc = 'R' THEN '9'+cPlazo ELSE '0'+ cPlazo END + " _
                 & "  CASE WHEN cProdCod IN ('305','320','423') THEN '01' ELSE RIGHT(cProdCod,2) END + " _
                 & "  CASE WHEN cFuenteFinanc = 'A' THEN '10' WHEN cFuenteFinanc = 'B' THEN '11' WHEN cFuenteFinanc = 'C' THEN '12' ELSE '0'+cFuenteFinanc END + " _
                 & "  CASE WHEN cAgeCod in ('99') THEN '11' ELSE cAgeCod END cCtaContCod, " _
                 & "  nSaldo " _
                 & "  FROM " & sServidorConsolidada & "CredSaldosCont WHERE cEstado = '1' and cRefinanc = 'N' and dFecha = '" & psFecha & "' and cConcepto = 'I' and cMoneda = '" & pnMoneda & "' ) Cred FULL OUTER JOIN " _
                 & "( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) nCtaSaldoImporte " _
                 & "  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
                 & "  (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' ) " _
                 & ") cs on cs.cCtaContCod = Cred.cCtaContCod " _
                 & "WHERE ISNULL(nSaldo,0) <> ISNULL(nCtaSaldoImporte,0) GROUP BY cred.cCtaContCod, cs.cCtaContCod, cs.nCtaSaldoImporte ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod) "
            End If
        End If
    End If
       
   If pnTipoAsiento = 2 Then  'En suspenso
        If pnBitCentral = True Then
             If sCodEmp = "112" Or sCodEmp = "108" Or sCodEmp = "109" Then
                 '** Final
                 'psSql = "SELECT cred.cCtaContCod cta1, SUM(ISNULL(cred.nSaldo,0)) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
                     & " FROM ( select '" & psCtaDebe & "0' " _
                     & "  + CASE WHEN cRefinanc = 'R' and cEstado IN ('1','3') THEN '1' WHEN cEstado = '5' THEN '2' WHEN cEstado = '6' THEN '3' ELSE cEstado END " _
                     & "  + '0'+left(cProdCod,1)  " _
                     & "  + CASE WHEN cProdCod = '" & Producto.gColConsuPrendario & "' THEN '13' " _
                     & "         WHEN cProdCod = '" & Producto.gColConsuPrestAdm & "' THEN '20' " _
                     & "         ELSE CASE WHEN cProdCod = '" & Producto.gColHipoMiVivienda & "' THEN '23' " _
                     & "                   WHEN cRefinanc = 'R' and cEstado IN ('5','6') THEN '19' " _
                     & "                     + CASE WHEN LEFT(cProdCod,1) IN ('1','2') THEN '29' else '' END + '06' " _
                     & "              ELSE '06' END " _
                     & "              + '0'+ cPlazo  + (SELECT LTRIM(RTRIM(cEquivAsiento)) From AsientoProdEquiv AE Where cEquivTpo ='SC' and AE.cProducto = cProdCod) " _
                     & "              +  cFuenteFinanc END" _
                     & "  +  cAgeCod as cCtaContCod, " _
                     & "  nSaldo " _
                     & "  FROM " & sServidorConsolidada & "CredSaldosCont WHERE ((cEstado = '1' and cRefinanc = 'R') OR cEstado <> '1') and dFecha = '" & psFecha & "' and cConcepto = 'I' and cMoneda = '" & pnMoneda & "' ) Cred " _
                     & "    FULL OUTER JOIN " _
                     & "    ( SELECT * FROM " _
                     & "        (   SELECT  cCtaContCod, CASE WHEN SubString(cCtaContCod,3,1) = '2' THEN ROUND(nCtaSaldoImporte / " & pnTipCambio & ",2) ELSE nCtaSaldoImporte END  nCtaSaldoImporte " _
                     & "            FROM    ctasaldo cs " _
                     & "            WHERE   cs.cCtaContCod like '" & psCtaDebe & "%' and not cs.cCtaContCod like '" & psCtaDebe & "090[349]%' and cs.dCtaSaldoFecha = " _
                     & "                    (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' ) " _
                     & "         ) AS X  WHERE ISNULL(X.nCtaSaldoImporte,0)<>0 " _
                     & "    ) cs on cs.cCtaContCod = Cred.cCtaContCod " _
                     & "GROUP BY cred.cCtaContCod, cs.cCtaContCod, cs.nCtaSaldoImporte ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"

                'CAMBIO EJRS 09 DE JUNIO 2005
                
                'By capi 09012009 para las cuentas contables producto de la reclasificacion de intereses devengados a suspenso por alineacion
                
                'psSql = "SELECT cred.cCtaContCod cta1, SUM(ISNULL(cred.nSaldo,0)) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
                '     & " FROM ( select '" & psCtaDebe & "0' " _
                '     & "  + CASE    WHEN cRefinanc = 'R' and cEstado IN ('1','3') THEN '1' " _
                '     & "            WHEN cEstado = '5' THEN '2'   " _
                '     & "            WHEN cEstado = '6' THEN '3' ELSE cEstado END " _
                '     & "            + '0'+left(cProdCod,1)    + " _
                '     & "            CASE    WHEN cRefinanc = 'R' and cEstado IN ('1','3') THEN '06' " _
                '     & "                    WHEN cRefinanc = 'N' AND cEstado = '5' AND LEFT(cProdCod,1)<>'3' THEN '06'   " _
                '     & "                    WHEN cRefinanc = 'R' AND cEstado = '5' AND LEFT(cProdCod,1)<>'3' THEN '06' " _
                '     & "                    WHEN cRefinanc = 'N' AND cEstado = '6' AND LEFT(cProdCod,1)<>'3' THEN '06'   " _
                '     & "                    WHEN cRefinanc = 'R' AND cEstado = '6' AND LEFT(cProdCod,1)<>'3' THEN '06' " _
                '     & "                    WHEN cRefinanc = 'R' AND LEFT(cProdCod,1)='3' AND cProdCod<>'320' THEN '06' " _
                '     & "                    WHEN cProdCod = '305' THEN '13' " _
                '     & "                    WHEN cProdCod = '320' THEN '2009' " _
                '     & "                    WHEN cProdCod = '423' THEN '23' " _
                '     & "                    Else '06' END + " _
                '     & "            CASE    WHEN cProdCod In ('305','320') THEN '' " _
                '     & "                    ELSE + '' + "
                     
                     
''     ALPA 20100625 B2*****************************************************************
''                psSql = "SELECT cred.cCtaContCod cta1, SUM(ISNULL(cred.nSaldo,0)) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
''                     & " FROM ( select '" & psCtaDebe & "0' " _
''                     & "  + CASE    WHEN cRefinanc = 'R' and cEstado IN ('1','3') THEN '1' " _
''                     & "            WHEN cEstado = '5' THEN '2'   " _
''                     & "            WHEN cEstado = '6' THEN '3'   " _
''                     & "            WHEN cEstado= '7' THEN '9' ELSE cEstado END " _
''                     & "            + '0'+left(cProdCod,1)    + " _
''                     & "            CASE    WHEN cRefinanc = 'R' and cEstado IN ('1','3') THEN '06' " _
''                     & "                    WHEN cRefinanc = 'N' AND cEstado = '5' AND LEFT(cProdCod,1)<>'3' THEN '06'   " _
''                     & "                    WHEN cRefinanc = 'R' AND cEstado = '5' AND LEFT(cProdCod,1)<>'3' THEN '06' " _
''                     & "                    WHEN cRefinanc = 'N' AND cEstado = '6' AND LEFT(cProdCod,1)<>'3' THEN '06'   " _
''                     & "                    WHEN cRefinanc = 'R' AND cEstado = '6' AND LEFT(cProdCod,1)<>'3' THEN '06' " _
''                     & "                    WHEN cRefinanc = 'R' AND LEFT(cProdCod,1)='3' AND cProdCod<>'320' THEN '06' " _
''                     & "                    WHEN cProdCod = '305' THEN '13' " _
''                     & "                    WHEN cProdCod = '320' THEN '2009' " _
''                     & "                    WHEN cProdCod = '423' THEN '23' " _
''                     & "                    Else '06' END + " _
''                     & "            CASE    WHEN cProdCod In ('305','320') THEN '' " _
''                     & "                    ELSE + '' + "
''
''                ' end by capi 09012009
''                      'ELSE + '0'+ cPlazo  +
''                psSql = psSql & "           ( SELECT    LTRIM(RTRIM(cEquivAsiento)) " _
''                     & "                                From AsientoProdEquiv AE " _
''                     & "                                Where cEquivTpo ='SC' and AE.cProducto = cProdCod) " _
''                     & "                    + (Select cCtaCont From dbconsolidada..coloclineacreditoequiv where clineacred = cFuenteFinanc) END  + cAgeCod as cCtaContCod , " _
''                     & "  nSaldo " _
''                     & "  FROM " & sServidorConsolidada & "CredSaldosCont WHERE ((cEstado = '1' and cRefinanc = 'R') OR cEstado <> '1') and dFecha = '" & psFecha & "' and cConcepto = 'I' and cMoneda = '" & pnMoneda & "' ) Cred " _
''                     & "    FULL OUTER JOIN " _
''                     & "    ( SELECT * FROM " _
''                     & "        (   SELECT  cCtaContCod, CASE WHEN SubString(cCtaContCod,3,1) = '2' THEN ROUND(nCtaSaldoImporte / " & pnTipCambio & ",2) ELSE nCtaSaldoImporte END  nCtaSaldoImporte " _
''                     & "            FROM    ctasaldo cs " _
''                     & "            WHERE   cs.cCtaContCod like '" & psCtaDebe & "%' and not cs.cCtaContCod like '" & psCtaDebe & "090[349]%' and cs.dCtaSaldoFecha = " _
''                     & "                    (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' ) " _
''                     & "         ) AS X  WHERE ISNULL(X.nCtaSaldoImporte,0)<>0 " _
''                     & "    ) cs on cs.cCtaContCod = Cred.cCtaContCod " _
''                     & "GROUP BY cred.cCtaContCod, cs.cCtaContCod, cs.nCtaSaldoImporte ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
psSql = "Exec B2_stp_sel_AjusteInteresesSuspenso '81', '" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "'," & pnTipCambio
''*************************************************************************************************
             End If
             If sCodEmp = "102" Then
                 '**Final
                 'psSql = "SELECT cred.cCtaContCod cta1, SUM(ISNULL(cred.nSaldo,0)) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
                     & "  FROM ( select '" & psCtaDebe & "0'+" _
                     & "  CASE WHEN cRefinanc = 'R' and cEstado IN ('1','3') THEN '1' WHEN cEstado = '5' THEN '2' WHEN cEstado = '6' THEN '3' ELSE cEstado END + " _
                     & "  '0'+left(cProdCod,1) + CASE WHEN cRefinanc = 'R' and not cEstado IN ('1','3') THEN '19' ELSE '' END + CASE WHEN cProdCod = '305' THEN '13' ELSE ISNULL((SELECT cEquivAsiento FROM AsientoProdEquiv eq WHERE eq.cProducto = cProdCod and cEquivTpo = 'CD'),'') END + " _
                     & "      ISNULL((SELECT cEquivAsiento FROM AsientoProdEquiv eq WHERE eq.cProducto = cProdCod and cEquivTpo = 'SC'),'') " _
                     & "  cCtaContCod, " _
                     & "  CASE WHEN cMoneda = '2' THEN ROUND(nSaldo * " & pnTipCambio & ",2) ELSE nSaldo END nSaldo " _
                     & "  FROM " & sServidorConsolidada & "CredSaldosCont WHERE ((cEstado = '1' and cRefinanc = 'R') OR cEstado <> '1') and dFecha = '" & psFecha & "' and cConcepto = 'I' and cMoneda = '" & pnMoneda & "' ) Cred FULL OUTER JOIN " _
                     & "( SELECT cCtaContCod, nCtaSaldoImporte " _
                     & "  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
                     & "  (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' ) " _
                     & ") cs on cs.cCtaContCod = Cred.cCtaContCod " _
                     & "GROUP BY cred.cCtaContCod, cs.cCtaContCod, cs.nCtaSaldoImporte ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
        
             End If
        ElseIf pnBitCentral = False Then
            If sCodEmp = "112" Then
                'psSql = "SELECT cred.cCtaContCod cta1, SUM(ISNULL(cred.nSaldo,0)) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte FROM ( select '" & psCtaDebe & "0'+" _
                    & "  CASE WHEN cRefinanc = 'R' and cEstado IN ('1','3') THEN '1' WHEN cEstado = '5' THEN '2' WHEN cEstado = '6' THEN '3' ELSE cEstado END + " _
                    & "  '0'+left(cProdCod,1) + CASE WHEN cProdCod = '305' THEN '13' WHEN cProdCod = '320' THEN '20' WHEN cProdCod = '423' THEN '23' WHEN cRefinanc = 'R' and cEstado IN ('5','6') THEN '19' Else '06' END + " _
                    & "  CASE WHEN cEstado IN ('5') and cDemandado='S' and LEFT(cProdCod,1) <> '3' and cRefinanc = 'R' THEN '29' WHEN cEstado IN ('5') and cDemandado = 'S' and LEFT(cProdCod,1) =  '3' THEN '9'+cPlazo WHEN cEstado IN ('5') and cDemandado = 'S' and not cRefinanc = 'R' THEN '9'+cPlazo ELSE '0'+ cPlazo END + " _
                    & "  CASE WHEN cProdCod IN ('305','320','423') THEN '01' ELSE RIGHT(cProdCod,2) END + " _
                    & "  CASE WHEN cFuenteFinanc = 'A' THEN '10' WHEN cFuenteFinanc = 'B' THEN '11' WHEN cFuenteFinanc = 'C' THEN '12' ELSE '0'+cFuenteFinanc END + " _
                    & "  CASE WHEN cAgeCod in ('99') THEN '11' ELSE cAgeCod END cCtaContCod, " _
                    & "  nSaldo " _
                    & "  FROM " & sServidorConsolidada & "CredSaldosCont WHERE ((cEstado = '1' and cRefinanc = 'R') OR not cEstado in ('1') ) and dFecha = '" & psFecha & "' and cConcepto = 'I' and cMoneda = '" & pnMoneda & "' ) Cred FULL OUTER JOIN " _
                    & "( SELECT cCtaContCod, CASE WHEN substring(cs.cCtaContCod, 3,1) = 2 THEN ROUND(nCtaSaldoImporte / " & pnTipCambio & ",2) ELSE nCtaSaldoImporte END nCtaSaldoImporte" _
                    & "  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and not cs.cCtaContCod like '" & psCtaDebe & "090[349]%' and cs.dCtaSaldoFecha = " _
                    & "  (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' ) " _
                    & ") cs on cs.cCtaContCod = Cred.cCtaContCod " _
                    & "GROUP BY cred.cCtaContCod, cs.cCtaContCod, cs.nCtaSaldoImporte ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
            End If
        End If
   End If
   
   If pnTipoAsiento = 3 Then  'Provisiones de Cartera
        If pnBitCentral = True Then
        
            If sCodEmp = "112" Or sCodEmp = "108" Or sCodEmp = "109" Then 'TRUJILLO /ICA
                '** Final
                ''ALPA 20081223*************************************************************************************
'                psSql = " SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, " _
'                      & " ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
'                      & " FROM ( Select '" & psClase & "' + '0'+SubString(cCtaCod,6,1) " _
'                      & "        + CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END " _
'                      & "        + CASE WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrendario & "' THEN '13' " _
'                      & "          WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColHipoMiVivienda & "' THEN '23' WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrestAdm & "' THEN '2009' " _
'                      & "          Else '06' END +  " _
'                      & "          Case  When SubString(c.cCtaCod,6,3) In ('305','320') Then '' Else '01' + ltrim(rtrim(CASE WHEN SubString(cCtaCod,6,3) IN ('" & Producto.gColConsuPrendario & "') THEN '01' " _
'                      & "          ELSE cEquivAsiento END)) +  E.cctacont End " _
'                      & "        + Substring(c.cCtaCod,4,2) cCtaContCod , " _
'                      & " SUM( nProvision ) nSaldo " _
'                      & " FROM ColocCalifProv c join " _
'                      & " (Select distinct substring(E.cLineaCred,1,4) Lin, cCtacont from " & sServidorConsolidada & "ColocLineaCreditoEquiv E ) E " _
'                      & "   on substring(c.cLineaCred,1,4)  = lin " _
'                      & " Join (SELECT LTRIM(RTRIM(cEquivAsiento)) cEquivAsiento, AE.cProducto From AsientoProdEquiv AE Where cEquivTpo ='SC') AE On AE.cProducto = SubString(c.cCtaCod,6,3)"
'
'                psSql = psSql & "  WHERE SubString(cCtaCod,9,1) = '" & pnMoneda & "' and nprdEstado in ('2020', '2021', '2022', '2030', '2031', '2032','2101', '2104', '2106', '2107', '2201','2205')  And substring(cctacod,7,2) <> '21' " _
'                      & " Group By '" & psClase & "' " _
'                      & "   + '0'+SubString(cCtaCod,6,1) " _
'                      & "   + CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END " _
'                      & "   + CASE WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrendario & "' THEN '13' " _
'                      & "          WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColHipoMiVivienda & "' THEN '23' WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrestAdm & "' THEN '2009' " _
'                      & "      Else '06' END + Case  When SubString(c.cCtaCod,6,3) In ('305','320') Then '' Else  '01' " _
'                      & "   + ltrim(rtrim(CASE WHEN SubString(cCtaCod,6,3) IN ('" & Producto.gColConsuPrendario & "') THEN '01' " _
'                      & "     ELSE cEquivAsiento END)) +  E.cctacont end " _
'                      & "   + Substring(c.cCtaCod,4,2)"
'
'                psSql = psSql & " UNION Select '" & psClase & "' + '0' + SubString(cCtaCod,6,1) " _
'                      & "        + '05' " _
'                      & "        + CASE WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrendario & "' THEN '13' " _
'                      & "          WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColHipoMiVivienda & "' THEN '23' WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrestAdm & "' THEN '2009' " _
'                      & "          Else '06' END + Case When SubString(c.cCtaCod,6,3) In ('305','320') Then '' Else  '01' " _
'                      & "        + ltrim(rtrim(CASE WHEN SubString(cCtaCod,6,3) IN ('" & Producto.gColConsuPrendario & "') THEN '01' " _
'                      & "          ELSE cEquivAsiento END)) +  E.cctacont End " _
'                      & "        + Substring(c.cCtaCod,4,2) cCtaContCod , " _
'                      & " SUM( nProvisionRCC  ) nSaldo " _
'                      & " FROM ColocCalifProv c join " _
'                      & " (Select distinct substring(E.cLineaCred,1,4) Lin, cCtacont from " & sServidorConsolidada & "ColocLineaCreditoEquiv E ) E " _
'                      & "   on substring(c.cLineaCred,1,4)  = lin " _
'                      & " Join (SELECT LTRIM(RTRIM(cEquivAsiento)) cEquivAsiento, AE.cProducto From AsientoProdEquiv AE Where cEquivTpo ='SC') AE On AE.cProducto = SubString(c.cCtaCod,6,3)"
'
'                psSql = psSql & "  WHERE SubString(cCtaCod,9,1) = '" & pnMoneda & "' and cCalGen = '0'  and nrcambiario = 2 And '" & pnMoneda & "' = '2' " _
'                      & " Group By '" & psClase & "' " _
'                      & "   + '0'+ SubString(cCtaCod,6,1) " _
'                      & "   + '05' " _
'                      & "   + CASE WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrendario & "' THEN '13' " _
'                      & "          WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColHipoMiVivienda & "' THEN '23' WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrestAdm & "' THEN '2009' " _
'                      & "      Else '06' END + Case  When SubString(c.cCtaCod,6,3) In ('305','320') Then '' Else '01' " _
'                      & "   + ltrim(rtrim(CASE WHEN SubString(cCtaCod,6,3) IN ('" & Producto.gColConsuPrendario & "') THEN '01' " _
'                      & "     ELSE cEquivAsiento END)) +  E.cctacont End " _
'                      & "   + Substring(c.cCtaCod,4,2)"
'
'                psSql = psSql & " " _
'                      & " ) Cred FULL OUTER JOIN " _
'                      & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) * -1 nCtaSaldoImporte  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
'                      & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' " _
'                      & "   ) ) cs on cs.cCtaContCod = Cred.cCtaContCod" _
'                      & " ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
                'ALP 20091006***************************
'                psSql = " SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, "
                    'IsNull(cred.cCtaContcod, cs.cCtaContCod)
                'ALPA 20100618 B2****************************************
                'nSaldoTotalProvProc = VerificarSumaProvisionProciclica
                '********************************************************
'                psSql = " SELECT cred.cCtaContcod cta1, case substring(IsNull(cred.cCtaContcod, cs.cCtaContCod),9,2) when '02' then " & IIf(nSaldoTotalProvProc > 0, "Round(ISNULL(cred.nSaldo,0),3)", "ISNULL(cs.nCtaSaldoImporte,0)") & " else Round(ISNULL(cred.nSaldo,0),3) end nSaldo, cs.cCtaContCod cta2, "
'                '****************************************
'                psSql = psSql & " ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte "
'                psSql = psSql & " FROM ( "
'
'                'ALPA 20090109*********************************************************************************
'                psSql = psSql & "  Select '" & psClase & "' + '0'+SubString(cCtaCod,6,1) "
'                psSql = psSql & "        + CASE WHEN cCalGen = 0 THEN '0201' ELSE '01' END "
'                psSql = psSql & "        + CASE WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrendario & "' THEN '13' "
'                psSql = psSql & "          WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColHipoMiVivienda & "' THEN '23' WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrestAdm & "' THEN '2009' "
'                psSql = psSql & "          Else '06' END +  "
'                psSql = psSql & "          Case  When SubString(c.cCtaCod,6,3) In ('305','320') Then '' Else '01' + ltrim(rtrim(CASE WHEN SubString(cCtaCod,6,3) IN ('" & Producto.gColConsuPrendario & "') THEN '01' "
'                psSql = psSql & "          ELSE cEquivAsiento END)) +  E.cctacont End "
'                psSql = psSql & "        + Substring(c.cCtaCod,4,2) cCtaContCod , "
'                psSql = psSql & " SUM( isnull(nProvision,0)) nSaldo "
'                psSql = psSql & " FROM ColocCalifProv c join "
'                psSql = psSql & " (Select distinct substring(E.cLineaCred,1,4) Lin, cCtacont from " & sServidorConsolidada & "ColocLineaCreditoEquiv E ) E "
'                psSql = psSql & "   on substring(c.cLineaCred,1,4)  = lin "
'                psSql = psSql & " Join (SELECT LTRIM(RTRIM(cEquivAsiento)) cEquivAsiento, AE.cProducto From AsientoProdEquiv AE Where cEquivTpo ='SC') AE On AE.cProducto = SubString(c.cCtaCod,6,3)"
'
'                psSql = psSql & "  WHERE SubString(cCtaCod,9,1) = '" & pnMoneda & "' and nprdEstado in ('2020', '2021', '2022', '2030', '2031', '2032','2101', '2104', '2106', '2107', '2201','2205')  And substring(cctacod,7,2) <> '21' "
'                psSql = psSql & " Group By '" & psClase & "' "
'                psSql = psSql & "   + '0'+SubString(cCtaCod,6,1) "
'                psSql = psSql & "   + CASE WHEN cCalGen = 0 THEN '0201' ELSE '01' END "
'                psSql = psSql & "   + CASE WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrendario & "' THEN '13' "
'                psSql = psSql & "          WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColHipoMiVivienda & "' THEN '23' WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrestAdm & "' THEN '2009' "
'                psSql = psSql & "      Else '06' END + Case  When SubString(c.cCtaCod,6,3) In ('305','320') Then '' Else  '01' "
'                psSql = psSql & "   + ltrim(rtrim(CASE WHEN SubString(cCtaCod,6,3) IN ('" & Producto.gColConsuPrendario & "') THEN '01' "
'                psSql = psSql & "     ELSE cEquivAsiento END)) +  E.cctacont end "
'                psSql = psSql & "   + Substring(c.cCtaCod,4,2)"
'                psSql = psSql & " UNION "
'                psSql = psSql & "  Select '" & psClase & "' + '0'+SubString(cCtaCod,6,1) "
'                psSql = psSql & "        + CASE WHEN cCalGen = 0 THEN '0202' ELSE '01' END "
'                psSql = psSql & "        + CASE WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrendario & "' THEN '13' "
'                psSql = psSql & "          WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColHipoMiVivienda & "' THEN '23' WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrestAdm & "' THEN '2009' "
'                psSql = psSql & "          Else '06' END +  "
'                psSql = psSql & "          Case  When SubString(c.cCtaCod,6,3) In ('305','320') Then '' Else '01' + ltrim(rtrim(CASE WHEN SubString(cCtaCod,6,3) IN ('" & Producto.gColConsuPrendario & "') THEN '01' "
'                psSql = psSql & "          ELSE cEquivAsiento END)) +  E.cctacont End "
'                psSql = psSql & "        + Substring(c.cCtaCod,4,2) cCtaContCod , "
'                psSql = psSql & " SUM( isnull(nProvisionProciclica,0)) nSaldo "
'                psSql = psSql & " FROM ColocCalifProv c join "
'                psSql = psSql & " (Select distinct substring(E.cLineaCred,1,4) Lin, cCtacont from " & sServidorConsolidada & "ColocLineaCreditoEquiv E ) E "
'                psSql = psSql & "   on substring(c.cLineaCred,1,4)  = lin "
'                psSql = psSql & " Join (SELECT LTRIM(RTRIM(cEquivAsiento)) cEquivAsiento, AE.cProducto From AsientoProdEquiv AE Where cEquivTpo ='SC') AE On AE.cProducto = SubString(c.cCtaCod,6,3)"
'
'                psSql = psSql & "  WHERE cCalGen = 0 and SubString(cCtaCod,9,1) = '" & pnMoneda & "' and nprdEstado in ('2020', '2021', '2022', '2030', '2031', '2032','2101', '2104', '2106', '2107', '2201','2205')  And substring(cctacod,7,2) <> '21' "
'                psSql = psSql & " Group By '" & psClase & "' "
'                psSql = psSql & "   + '0'+SubString(cCtaCod,6,1) "
'                psSql = psSql & "   + CASE WHEN cCalGen = 0 THEN '0202' ELSE '01' END "
'                psSql = psSql & "   + CASE WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrendario & "' THEN '13' "
'                psSql = psSql & "          WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColHipoMiVivienda & "' THEN '23' WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrestAdm & "' THEN '2009' "
'                psSql = psSql & "      Else '06' END + Case  When SubString(c.cCtaCod,6,3) In ('305','320') Then '' Else  '01' "
'                psSql = psSql & "   + ltrim(rtrim(CASE WHEN SubString(cCtaCod,6,3) IN ('" & Producto.gColConsuPrendario & "') THEN '01' "
'                psSql = psSql & "     ELSE cEquivAsiento END)) +  E.cctacont end "
'                psSql = psSql & "   + Substring(c.cCtaCod,4,2)"
'
'                psSql = psSql & " UNION "
'                psSql = psSql & " Select '" & psClase & "' + '0' + SubString(cCtaCod,6,1) "
'                psSql = psSql & "        + '05' "
'                psSql = psSql & "        + CASE WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrendario & "' THEN '13' "
'                psSql = psSql & "          WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColHipoMiVivienda & "' THEN '23' WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrestAdm & "' THEN '2009' "
'                psSql = psSql & "          Else '06' END + Case When SubString(c.cCtaCod,6,3) In ('305','320') Then '' Else  '01' "
'                psSql = psSql & "        + ltrim(rtrim(CASE WHEN SubString(cCtaCod,6,3) IN ('" & Producto.gColConsuPrendario & "') THEN '01' "
'                psSql = psSql & "          ELSE cEquivAsiento END)) +  E.cctacont End "
'                psSql = psSql & "        + Substring(c.cCtaCod,4,2) cCtaContCod , "
'                psSql = psSql & " SUM( nProvisionRCC  ) nSaldo "
'                psSql = psSql & " FROM ColocCalifProv c join "
'                psSql = psSql & " (Select distinct substring(E.cLineaCred,1,4) Lin, cCtacont from " & sServidorConsolidada & "ColocLineaCreditoEquiv E ) E "
'                psSql = psSql & "   on substring(c.cLineaCred,1,4)  = lin "
'                psSql = psSql & " Join (SELECT LTRIM(RTRIM(cEquivAsiento)) cEquivAsiento, AE.cProducto From AsientoProdEquiv AE Where cEquivTpo ='SC') AE On AE.cProducto = SubString(c.cCtaCod,6,3)"
'
'                psSql = psSql & "  WHERE SubString(cCtaCod,9,1) = '" & pnMoneda & "' and cCalGen = '0'  and nrcambiario = 2 And '" & pnMoneda & "' = '2' "
'                psSql = psSql & " Group By '" & psClase & "' "
'                psSql = psSql & "   + '0'+ SubString(cCtaCod,6,1) "
'                psSql = psSql & "   + '05' "
'                psSql = psSql & "   + CASE WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrendario & "' THEN '13' "
'                psSql = psSql & "          WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColHipoMiVivienda & "' THEN '23' WHEN SubString(c.cCtaCod,6,3)='" & Producto.gColConsuPrestAdm & "' THEN '2009' "
'                psSql = psSql & "      Else '06' END + Case  When SubString(c.cCtaCod,6,3) In ('305','320') Then '' Else '01' "
'                psSql = psSql & "   + ltrim(rtrim(CASE WHEN SubString(cCtaCod,6,3) IN ('" & Producto.gColConsuPrendario & "') THEN '01' "
'                psSql = psSql & "     ELSE cEquivAsiento END)) +  E.cctacont End "
'                psSql = psSql & "   + Substring(c.cCtaCod,4,2)"
'
'                psSql = psSql & " "
'                psSql = psSql & " ) Cred FULL OUTER JOIN "
'                psSql = psSql & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) * -1 nCtaSaldoImporte  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = "
'                psSql = psSql & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' "
'                psSql = psSql & "   ) ) cs on cs.cCtaContCod = Cred.cCtaContCod"
'                psSql = psSql & " ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
                
                psSql = "exec B2_stp_sel_AjusteProvisionCarteraCreditos '14','" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "'," & pnTipCambio & ",'" & DevuelveMesAnio(psFecha, 1) & "','" & DevuelveMesAnio(psFecha, 2) & "'," & nSaldoTotalProvProc
              End If
              If sCodEmp = "102" Then  'LIMA
                '** Final
                'psSql = "SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
                   & "FROM ( Select '" & psClase & "'+ " _
                   & "          '0'+SubString(cCtaCod,6,1) + " _
                   & "          CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END cCtaContCod , " _
                   & "          SUM( nProvision ) nSaldo " _
                   & "    FROM COlocCalifProv " _
                   & "    WHERE SubString(cCtaCod,9,1) = '" & pnMoneda & "' and nprdestado in ('" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "','" & gColPEstDesem & "', '" & gColPEstVenci & "', '" & gColPEstPRema & "', '" & gColPEstRenov & "', '" & gColocEstRecVigJud & "') " _
                   & "    Group By '" & psClase & "'+'0'+SubString(cCtaCod,6,1) + " _
                   & "          " _
                   & "          CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END" _
                   & "    ) Cred FULL OUTER JOIN " _
                   & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) * -1 nCtaSaldoImporte  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
                   & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' " _
                   & "   ) ) cs on cs.cCtaContCod = Cred.cCtaContCod " _
                   & "ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
              End If
        ElseIf pnBitCentral = False Then
            If sCodEmp = "112" Then  'TRUJILLO
              'psSql = "SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
                 & "FROM ( select '" & psClase & "'+ " _
                 & "          '0'+SubString(cCodCta,3,1) + " _
                 & "          CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END + " _
                 & "          CASE WHEN SubString(cCodCta,3,3)='305' THEN '13' WHEN SubString(cCodCta,3,3) = '320' THEN '20' WHEN SubString(cCodCta,3,3) = '423' THEN '23' Else '06' END + " _
                 & "          CASE WHEN SubString(cCodCta,3,3) IN ('305','320','423') THEN '01' ELSE SubString(cCodCta,4,2) END + " _
                 & "          CASE WHEN SubString(cCodLinCred,6,1) = 'A' THEN '10' WHEN SubString(cCodLinCred,6,1) = 'B' THEN '11' WHEN SubString(cCodLinCred,6,1) = 'C' THEN '12' ELSE '0'+SubString(cCodLinCred,6,1) END + " _
                 & "          CASE WHEN Left(cCodCta,2) in ('99') THEN '11' ELSE Left(cCodCta,2) END cCtaContCod , " _
                 & "          SUM( nProvision ) nSaldo " _
                 & "    FROM " & sServidorConsolidada & "CreditoAudi " _
                 & "    WHERE SubString(cCodCta,6,1) = '" & pnMoneda & "' and cEstado IN ('F','1','4','6','7','V') and not SubString(cCodCta,4,2) in ('21') " _
                 & "    Group By '" & psClase & "'+'0'+SubString(cCodCta,3,1) + " _
                 & "       CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END + " _
                 & "       CASE WHEN SubString(cCodCta,3,3)='305' THEN '13' WHEN SubString(cCodCta,3,3) = '320' THEN '20' WHEN SubString(cCodCta,3,3) = '423' THEN '23' Else '06' END + " _
                 & "       CASE WHEN SubString(cCodCta,3,3) IN ('305','320','423') THEN '01' ELSE SubString(cCodCta,4,2) END + " _
                 & "       CASE WHEN SubString(cCodLinCred,6,1) = 'A' THEN '10' WHEN SubString(cCodLinCred,6,1) = 'B' THEN '11' WHEN SubString(cCodLinCred,6,1) = 'C' THEN '12' ELSE '0'+SubString(cCodLinCred,6,1) END + " _
                 & "       CASE WHEN Left(cCodCta,2) in ('99') THEN '11' ELSE Left(cCodCta,2) END " _
                 & "    ) Cred FULL OUTER JOIN " _
                 & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) * -1 nCtaSaldoImporte  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
                 & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' " _
                 & "   ) ) cs on cs.cCtaContCod = Cred.cCtaContCod " _
                 & "ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
            End If
        End If
   End If

   If pnTipoAsiento = 5 Or pnTipoAsiento = 6 Or pnTipoAsiento = 22 Then  'Capital o Intereses de Creditos Castigados 'JUEZ 20130116 Se agregó pnTipoAsiento = 22 para Gastos de Creditos Castigados
        If pnBitCentral = True Then
            If sCodEmp = "112" Or sCodEmp = "108" Or sCodEmp = "109" Then 'TRUJILLO / ICA
                '**Comentado por DAOR 20080723 ****************************************************************************
                '  psSql = "SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
                '     & " FROM ( select '" & psClase & "' " _
                '     & "     + '0'+SubString(cCtaCod,6,1)  " _
                '     & "     + CASE WHEN SubString(cCtaCod,6,3) = '" & Producto.gColHipoMiVivienda & "' THEN '23' Else '06' END " _
                '     & "     + cEquivAsiento +  left(E.cCtaCont,2) " _
                '     & "     + Substring(cCtaCod,4,2) cCtaContCod, " _
                '     & "   " & IIf(pnTipoAsiento = 5, " SUM(nSaldoCap) ", "SUM(nIntDev) ") & " nSaldo " _
                '     & "    FROM " & sServidorConsolidada & "CreditoConsol c Join " _
                '    & "    (Select distinct substring(E.cLineaCred,1,4) Lin, cCtacont from " & sServidorConsolidada & "ColocLineaCreditoEquiv E ) E " _
                '     & "    On substring(c.cLineaCred,1,4)  = lin Join (SELECT LTRIM(RTRIM(cEquivAsiento)) cEquivAsiento, cProducto From AsientoProdEquiv AE Where cEquivTpo ='SC') AE On AE.cProducto = SubString(cCtaCod,6,3) " _
                '     & "  WHERE substring(C.cCtaCod,9,1) = '" & pnMoneda & "' and C.nPrdEstado IN('" & gColocEstRecVigCast & "','2206') and C.nSaldoCap >= 0 " _
                '     & "  Group By '" & psClase & "'+ '0'+SubString(cCtaCod,6,1) + " _
                '     & "     + CASE WHEN SubString(cCtaCod,6,3) = '" & Producto.gColHipoMiVivienda & "' THEN '23' Else '06' END " _
                '     & "     + cEquivAsiento +  left(E.cCtaCont,2)  " _
                '     & "     + Substring(cCtaCod,4,2)  " _
                '     & "    ) Cred FULL OUTER JOIN " _
                '     & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) nCtaSaldoImporte  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
                '     & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' " _
                '     & "   ) ) cs on cs.cCtaContCod = Cred.cCtaContCod " _
                '     & "ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
                '*********************************************************************************************************
                
                '**DAOR 20080723, De acuerdo a resolución SBS Nº 2032-2008 y Acta Nº 158-2008/TI-D ***********************
                Dim lsCtaCtbCastigAmor As String
                'JUEZ 20130116 *********************************
                'lsCtaCtbCastigAmor = IIf(pnTipoAsiento = 5, "81" & pnMoneda & "92501", "81" & pnMoneda & "92502")
                lsCtaCtbCastigAmor = "81" & pnMoneda & IIf(pnTipoAsiento = 5, "92501", IIf(pnTipoAsiento = 6, "92502", "92503"))
                'END JUEZ **************************************
                
'                psSql = "select Cred.cCtaContcod cta1, round(isnull(Cred.nSaldo,0),3) nSaldo, CS.cCtaContCod cta2, isnull(CS.nCtaSaldoImporte,0) nCtaSaldoImporte "
'                psSql = psSql & " from (select (case when PG.nCredEstado is null then '" & psClase & "' else '" & lsCtaCtbCastigAmor & "' end) + '0'+subString(C.cCtaCod,6,1) "
'                psSql = psSql & "           + (case when substring(C.cCtaCod,6,3) = '" & Producto.gColHipoMiVivienda & "' then '23' else '06' end) "
'                psSql = psSql & "           + cEquivAsiento +  left(E.cCtaCont,2) + substring(C.cCtaCod,4,2) cCtaContCod, "
'                psSql = psSql & "        " & IIf(pnTipoAsiento = 5, " sum(nSaldoCap) ", "sum(nIntDev) ") & " nSaldo "
'                psSql = psSql & "       from " & sServidorConsolidada & "CreditoConsol C "
'                psSql = psSql & "           inner join (select distinct substring(E.cLineaCred,1,4) Lin, cCtacont from " & sServidorConsolidada & "ColocLineaCreditoEquiv E) E "
'                psSql = psSql & "               on substring(C.cLineaCred,1,4)  = E.Lin "
'                psSql = psSql & "           inner join (select ltrim(rtrim(cEquivAsiento)) cEquivAsiento, cProducto from AsientoProdEquiv AE where cEquivTpo ='SC') AE "
'                psSql = psSql & "               on SubString(C.cCtaCod,6,3)=AE.cProducto "
'                psSql = psSql & "           left join (select distinct MC.cCtaCod,MC.nCredEstado from Mov M inner join MovCol MC on M.nMovNro=MC.nMovNro and M.nMovFlag=0 "
'                psSql = psSql & "                     where MC.nCredEstado in (2202,2206) and MC.cOpeCod like '130%' and left(M.cMovNro,8)<='" & Format(psFecha, "yyyymmdd") & "') PG "
'                psSql = psSql & "               on C.cCtaCod=PG.cCtaCod "
'                psSql = psSql & "       where substring(C.cCtaCod,9,1) = '" & pnMoneda & "' and C.nPrdEstado in('2202','2206') and C.nSaldoCap >= 0"
'                psSql = psSql & "       group by (case when PG.nCredEstado is null then '" & psClase & "' else '" & lsCtaCtbCastigAmor & "' end) + '0'+subString(C.cCtaCod,6,1) "
'                psSql = psSql & "           + (case when substring(C.cCtaCod,6,3) = '" & Producto.gColHipoMiVivienda & "' then '23' else '06' end) "
'                psSql = psSql & "           + cEquivAsiento +  left(E.cCtaCont,2) + substring(C.cCtaCod,4,2) "
'                psSql = psSql & "     ) Cred "
'                psSql = psSql & "     full outer join "
'                psSql = psSql & "     ( "
'                psSql = psSql & "     select cCtaContCod, round(nCtaSaldoImporte " & IIf(pnMoneda = 2, "/" & pnTipCambio, "") & ",2) nCtaSaldoImporte "
'                psSql = psSql & "     from CtaSaldo CS "
'                psSql = psSql & "     where (CS.cCtaContCod like '" & psCtaDebe & "%' or CS.cCtaContCod like '" & lsCtaCtbCastigAmor & "%') "
'                psSql = psSql & "         and CS.dCtaSaldoFecha = (select max(dCtaSaldoFecha) from CtaSaldo CS1 where CS1.cCtaContCod = CS.cCtaContCod and CS1.dCtaSaldoFecha <= '" & psFecha & "')"
'                psSql = psSql & "     ) CS on CS.cCtaContCod = Cred.cCtaContCod "
'                psSql = psSql & " order by isnull(CS.cCtaContCod,Cred.cCtaContCod) "
                '****************************************************************************************
                'JUEZ 20130116 **************************
                'psSql = "exec B2_stp_sel_AjusteCapitalCreditosCastigados '81', '" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "'," & pnTipCambio & ",'" & IIf(pnTipoAsiento = 5, "01", "02") & "'"
                psSql = "exec B2_stp_sel_AjusteCapitalCreditosCastigados '81', '" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "'," & pnTipCambio & ",'" & IIf(pnTipoAsiento = 5, "01", IIf(pnTipoAsiento = 6, "02", "03")) & "'"
                'END JUEZ *******************************
            End If
             
            If sCodEmp = "102" Then 'LIMA
                  'psSql = "SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
                     & "FROM ( " _
                     & "     select cCtaContCod, SUM(nSaldo) nSaldo FROM ( " _
                     & "         select '" & psClase & "'+ " _
                     & "        '0'+SubString(cCtaCod,6,1) + " _
                     & "        ISNULL(CASE WHEN SubString(cCtaCod,6,3) = '305' THEN '13' ELSE (SELECT cEquivAsiento FROM AsientoProdEquiv eq WHERE eq.cProducto = SubString(cCtaCod,6,3) and cEquivTpo = 'CD') END,'') + " _
                     & "        RTRIM(ISNULL((SELECT cEquivAsiento FROM AsientoProdEquiv eq WHERE eq.cProducto = SubString(cCtaCod,6,3) and cEquivTpo = 'SC'),'')) cCtaContCod, " _
                     & "        " & IIf(pnTipoAsiento = 5, " Round(nSaldoCap,2) ", "Round(nSaldIntcom,2) ") & " nSaldo " _
                     & "       FROM " & sServidorConsolidada & "CreditoConsol c " _
                     & "       WHERE substring(C.cCtaCod,9,1) = '" & pnMoneda & "' and C.nPrdEstado IN('" & gColocEstRecVigCast & "') and C.nSaldoCap > 0  " _
                     & "       ) AAA Group By cCtaContCod " _
                     & "    ) Cred FULL OUTER JOIN " _
                     & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) nCtaSaldoImporte  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
                     & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' " _
                     & "   ) ) cs on cs.cCtaContCod = Cred.cCtaContCod " _
                     & "ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
            End If
        ElseIf pnBitCentral = False Then
            'psSql = "SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
               & "FROM ( select '" & psClase & "'+ " _
               & "        '0'+SubString(cCodCta,3,1) + " _
               & "        CASE WHEN SubString(cCodCta,3,3)='305' THEN '13' WHEN SubString(cCodCta,3,3) = '320' THEN '20' WHEN SubString(cCodCta,3,3) = '423' THEN '23' Else '06' END + " _
               & "        '0'+ SubString(cCodLinCred,5,1) + " _
               & "        CASE WHEN SubString(cCodCta,3,3) IN ('305','320','423') THEN '01' ELSE SubString(cCodCta,4,2) END + " _
               & "        CASE WHEN SubString(cCodLinCred,6,1) = 'A' THEN '10' WHEN SubString(cCodLinCred,6,1) = 'B' THEN '11' WHEN SubString(cCodLinCred,6,1) = 'C' THEN '12' ELSE '0'+SubString(cCodLinCred,6,1) END + " _
               & "        CASE WHEN Left(cCodCta,2) in ('99') THEN '99' ELSE Left(cCodCta,2) END cCtaContCod, " _
               & "        " & IIf(pnTipoAsiento = 5, " SUM(nSaldoCap) ", "SUM(nSaldIntcom) ") & " nSaldo " _
               & "       FROM " & sServidorConsolidada & "CreditoConsol c " _
               & "       WHERE substring(C.cCodCta,6,1) = '" & pnMoneda & "' and C.cEstado IN ('V') and cCondCre = 'A' and C.nSaldoCap > 0 " _
               & "Group By '" & psClase & "'+ '0'+SubString(cCodCta,3,1) + " _
               & "        CASE WHEN SubString(cCodCta,3,3)='305' THEN '13' WHEN SubString(cCodCta,3,3) = '320' THEN '20' WHEN SubString(cCodCta,3,3) = '423' THEN '23' Else '06' END + " _
               & "        '0'+ SubString(cCodLinCred,5,1) + " _
               & "        CASE WHEN SubString(cCodCta,3,3) IN ('305','320','423') THEN '01' ELSE SubString(cCodCta,4,2) END + " _
               & "        CASE WHEN SubString(cCodLinCred,6,1) = 'A' THEN '10' WHEN SubString(cCodLinCred,6,1) = 'B' THEN '11' WHEN SubString(cCodLinCred,6,1) = 'C' THEN '12' ELSE '0'+SubString(cCodLinCred,6,1) END + " _
               & "        CASE WHEN Left(cCodCta,2) in ('99') THEN '99' ELSE Left(cCodCta,2) END " _
               & "    ) Cred FULL OUTER JOIN " _
               & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) nCtaSaldoImporte  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
               & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' " _
               & "   ) ) cs on cs.cCtaContCod = Cred.cCtaContCod " _
               & "ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
        End If
    End If
    
   If pnTipoAsiento = 7 Then 'Reversión de Intereses Devengados
        If pnBitCentral = True Then
            If sCodEmp = "112" Or sCodEmp = "108" Or sCodEmp = "109" Then 'Trujillo / ICA
             ' ** final
                If pbAsiRevDev = False Then
                    psSql = " SELECT '" & psClase & "' " _
                          & " + '0' + SubString(ca.cCtaCod,6,1) " _
                          & " + CASE WHEN SubString(ca.cCtaCod,6,3) = '" & Producto.gColConsuPrendario & "' THEN '13' WHEN SubString(ca.cCtaCod,6,3) = '" & Producto.gColConsuPrestAdm & "' THEN '2009' Else '06' +  SC.SubTipo + left(E.cCtaCont,2) End " _
                          & " + Substring(ca.cCtaCod,4,2)  Cta1,  " _
                          & "  sum(cc.nIntDev) nSaldo, NULL Cta2, 0 nCtaSaldoImporte " _
                          & "  FROM ColocCalifProv ca " _
                          & "  JOIN " & sServidorConsolidada & "CreditoConsol cc ON CC.cCtaCod = ca.cCtaCod " _
                          & "  JOIN (Select distinct substring(E.cLineaCred,1,4) Lin, cCtacont from " & sServidorConsolidada & "ColocLineaCreditoEquiv E ) E " _
                          & "     On substring(ca.cLineaCred,1,4)  = lin " _
                          & "  JOIN (SELECT distinct cProducto, LTRIM(RTRIM(cEquivAsiento)) SubTipo From AsientoProdEquiv AE Where cEquivTpo ='SC'  ) SC " _
                          & "     On SC.cProducto = substring(ca.cCtaCod,6,3) " _
                          & "  WHERE ca.cCtaCod LIKE '________" & pnMoneda & "%' and ca.cRefinan = 'N' and ca.nPrdEstado in ('" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "' ,'" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColPEstDesem & "', '" & gColPEstVenci & "', '" & gColPEstPRema & "', '" & gColPEstRenov & "')  and CC.nPrdEstado in ('" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "' ,'" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColPEstDesem & "', '" & gColPEstVenci & "', '" & gColPEstPRema & "', '" & gColPEstRenov & "' ) " _
                          & "     and ( (SubString(ca.cCtaCod,6,1) = 1 and ca.nDiasAtraso <= 15) or " _
                          & "           (SubString(ca.cCtaCod,6,1) in (2,3,4) and ca.nDiasAtraso <= 30)  ) " _
                          & "     and ca.cCalGen in (3,4) " _
                          & "  group By '" & psClase & "' " _
                          & " + '0' + SubString(ca.cCtaCod,6,1)  " _
                          & " + CASE WHEN SubString(ca.cCtaCod,6,3) = '" & Producto.gColConsuPrendario & "' THEN '13' Else '06' +  SC.SubTipo + left(E.cCtaCont,2) End  " _
                          & " + Substring(ca.cCtaCod,4,2) "
                Else
'                    psSql = " SELECT  Cta1, cCtaContCodREV as Cta2, cCtaSuspDebe as Cta3, cCtaSuspHaber as Cta4, SUM(nSaldo) as nSaldo, 0 as nCtaSaldoImporte "
'                    psSql = psSql & " From "
'                    psSql = psSql & "         ("
'                    psSql = psSql & "           SELECT '51' + SubString(cCtaCod,9,1) + '401'+'0'+SubString(cCtaCod,6,1) + "
'                    '**Modificado por DAOR 20080308, Por Observación de Contabilidad LEMA *******
'                    'psSql = psSql & "                   CASE WHEN SubString(cCtaCod,6,3) = '305' THEN '13' WHEN SubString(cCtaCod,6,3) = '320' THEN '2009' WHEN SubString(cCtaCod,6,3) = '423' THEN '23' Else '06' End + Case When SubString(cCtaCod,6,3) In ('305','320') Then '' When SubString(cCtaCod,6,1) = '3' Then '' Else '' End +   (SELECT LTRIM(RTRIM(cEquivAsiento)) From AsientoProdEquiv AE Where cEquivTpo ='SC' and AE.cProducto = cProdCod) + Case When SubString(cCtaCod,6,3) In ('305','320') then '' else cFuenteFinanc end "
'                    psSql = psSql & "                   CASE WHEN SubString(cCtaCod,6,3) = '305' THEN '13' WHEN SubString(cCtaCod,6,3) = '320' THEN '2009' WHEN SubString(cCtaCod,6,3) = '423' THEN '23' Else '06' End + Case When SubString(cCtaCod,6,3) In ('305','320') Then '' When SubString(cCtaCod,6,1) = '3' Then '' Else '' End +  case when cProdCod <>'320' then (SELECT LTRIM(RTRIM(cEquivAsiento)) From AsientoProdEquiv AE Where cEquivTpo ='SC' and AE.cProducto = cProdCod) else '' end + Case When SubString(cCtaCod,6,3) In ('305','320') then '' else cFuenteFinanc end "
'                    '****************************************************************************
'                    psSql = psSql & "                   + Substring(cCtaCod,4,2) Cta1,"
'                    psSql = psSql & "                   '14'+ cMoneda +'8' + '0'+left(cProdCod,1)  +"
'                    psSql = psSql & "                           CASE    WHEN cProdCod = '320' THEN '2009'"
'                    psSql = psSql & "                                   WHEN cProdCod = '423' THEN '23' WHEN cProdCod = '305' THEN '13' Else '06' END + Case When cProdCod In ('305','320') Then '' Else '0'+ cPlazo +"
'                    psSql = psSql & "                           CASE    WHEN cProdCod ='305' THEN '01'"
'                    psSql = psSql & "                                   ELSE (  SELECT LTRIM(RTRIM(cEquivAsiento))"
'                    psSql = psSql & "                                           From AsientoProdEquiv AE"
'                    psSql = psSql & "                                           Where cEquivTpo ='SC' and AE.cProducto = cProdCod) End + cFuenteFinanc End + "
'                    psSql = psSql & "                                           cAgeCod  as cCtaContCodREV ,"
'                    psSql = psSql & "                       nIntDev nSaldo, "
'                    psSql = psSql & "                   '81' + cMoneda + '409' + CASE cCalGen when '3' then '01' else '02' end + '0'+ substring(cCtaCod,6,1) + CASE    WHEN cProdCod In ('305','320') THEN '01' ELSE (SELECT LTRIM(RTRIM(cEquivAsiento)) From AsientoProdEquiv AE Where cEquivTpo ='SC' and AE.cProducto = cProdCod) End + cFuenteFinanc + cAgeCod as cCtaSuspDebe, "
'                    psSql = psSql & "                   '82'+ cMoneda + '1' as cCtaSuspHaber  "
'                    psSql = psSql & "            From"
'                    psSql = psSql & "               ("
'                    psSql = psSql & "                   SELECT  ca.cCtaCod, cc.nIntDev, SC.SUBTIPO,E.cCtacont, ca.cCalGen, "
'                    psSql = psSql & "                           SUBSTRING(cc.CCTACOD,9,1) AS cMoneda, cc.cReFinan as  cReFinanc,"
'                    psSql = psSql & "                           SUBSTRING(cc.CCTACOD,6,3) as cProdCod, SUBSTRING(cc.cLineaCred,6,1) as cPlazo,"
'                    psSql = psSql & "                           cFuenteFinanc = (Select top 1 cctacont FFinan  From " & sServidorConsolidada & "ColocLineaCreditoEquiv Where substring(cLineaCred,1,4)=LEFT(CC.cLineaCred,4)),"
'                    psSql = psSql & "                           SUBSTRING(CC.cCtaCod,4,2) as cAgeCod,"
'                    psSql = psSql & "                           cEstado = CASE WHEN SUBSTRING(CC.CCTACOD,6,1)='1' AND CC.nDiasAtraso<=15 and CC.nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and CC.cRefinan ='N' then '1'"
'                    psSql = psSql & "                           WHEN SUBSTRING(CC.CCTACOD,6,1)='2' AND CC.nDiasAtraso<=30 and CC.nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and CC.cRefinan ='N' then '1'"
'                    psSql = psSql & "                           WHEN SUBSTRING(CC.CCTACOD,6,1)='3' AND CC.nDiasAtraso<=30 and CC.nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and CC.cRefinan ='N' then '1'"
'                    psSql = psSql & "                           WHEN SUBSTRING(CC.CCTACOD,6,1)='4' AND CC.nDiasAtraso<=30 and CC.nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and CC.cRefinan ='N' then '1'"
'                    psSql = psSql & "                           Else '*' END"
'                    psSql = psSql & "                    FROM    ColocCalifProv ca"
'                    psSql = psSql & "                            JOIN " & sServidorConsolidada & "CreditoConsol cc ON CC.cCtaCod = ca.cCtaCod"
'                    psSql = psSql & "                            JOIN (     Select  distinct substring(E.cLineaCred,1,4) Lin, cCtacont"
'                    psSql = psSql & "                                       from    " & sServidorConsolidada & "ColocLineaCreditoEquiv E ) E"
'                    psSql = psSql & "                                               On substring(ca.cLineaCred,1,4)  = lin"
'                    psSql = psSql & "                             JOIN (  SELECT distinct cProducto, LTRIM(RTRIM(cEquivAsiento)) SubTipo"
'                    psSql = psSql & "                                       From AsientoProdEquiv AE"
'                    psSql = psSql & "                                       Where cEquivTpo ='SC'  ) SC"
'                    psSql = psSql & "                                       On SC.cProducto = substring(ca.cCtaCod,6,3)"
'                    psSql = psSql & "                    WHERE  ca.cCtaCod LIKE '________" & pnMoneda & "%' and ca.cRefinan = 'N'"
'                    psSql = psSql & "                           and ca.nPrdEstado in ('2020', '2021', '2022' ,'2030', '2031', '2032', '2101', '2104', '2106', '2107')"
'                    psSql = psSql & "                           and CC.nPrdEstado in ('2020', '2021', '2022' ,'2030', '2031', '2032', '2101', '2104', '2106', '2107' )"
'                    psSql = psSql & "                           and ( (SubString(ca.cCtaCod,6,1) = 1 and ca.nDiasAtraso <= 15) or (SubString(ca.cCtaCod,6,1) in (2,3,4) and ca.nDiasAtraso <= 30) )"
'                    psSql = psSql & "                           and ca.cCalGen in (3,4)"
'                    psSql = psSql & "                  ) AS RevInt"
'                    psSql = psSql & "          ) AS TOTAL"
'                    psSql = psSql & " GROUP BY Cta1, cCtaContCodREV, cCtaSuspDebe, cCtaSuspHaber"
                psSql = "exec B2_stp_sel_ReversionInteresesDevengados '51', '" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "'," & pnTipCambio
                End If
            End If
             
            If sCodEmp = "102" Then 'Lima
             ' ** final
             ' psSql = " SELECT '" & psClase & "'+ '0'+SubString(ca.cCtaCod,6,1) + CASE WHEN SubString(ca.cCtaCod,6,3) = '" & Producto.gColConsuPrendario & "' THEN '13' WHEN SubString(ca.cCtaCod,6,3) = '" & Producto.gColConsuPrestAdm & "' THEN '20' WHEN SubString(ca.cCtaCod,6,3) = '" & Producto.gColHipoMiVivienda & "' THEN '23' Else '06' END +" _
                    & "  '0'+ SubString(cc.cLineaCred,6,1) + " _
                    & "  CASE WHEN SubString(ca.cCtaCod,6,3) IN ('" & Producto.gColConsuPrendario & "','" & Producto.gColConsuPrestAdm & "','" & Producto.gColHipoMiVivienda & "') THEN '01' ELSE SubString(ca.cCtaCod,7,2) END + " _
                    & "  '0'+SubString(cc.cLineaCred,2,1) +  CASE WHEN SUBSTRING(ca.cCtaCod,4,2) in ('99') THEN '11' ELSE SUBSTRING(ca.cCtaCod,4,2) END Cta1, " _
                    & "  sum(cc.nIntDev) nSaldo, NULL Cta2, 0 nCtaSaldoImporte " _
                    & "  FROM ColocCalifProv ca JOIN " & sServidorConsolidada & "CreditoConsol cc ON CC.cCtaCod = ca.cCtaCod "
             
             'psSql = psSql & "  WHERE ca.cCtaCod LIKE '________" & pnMoneda & "%' and ca.cRefinan = 'N' and ca.nPrdEstado in ('" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "' ,'" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColPEstDesem & "', '" & gColPEstVenci & "', '" & gColPEstPRema & "', '" & gColPEstRenov & "')  and CC.nPrdEstado in ('" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "' ,'" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "', '" & gColPEstDesem & "', '" & gColPEstVenci & "', '" & gColPEstPRema & "', '" & gColPEstRenov & "' ) " _
                    & "     and ( (SubString(ca.cCtaCod,6,1) = 1 and ca.nDiasAtraso <= 15) or " _
                    & "      (SubString(ca.cCtaCod,6,1) in (2,3,4) and ca.nDiasAtraso <= 30)  ) " _
                    & "  and ca.cCalGen in (3,4) " _
                    & "  group By '" & psClase & "'+  '0'+SubString(ca.cCtaCod,6,1) + CASE WHEN SubString(ca.cCtaCod,6,3) = '" & Producto.gColConsuPrendario & "' THEN '13' WHEN SubString(ca.cCtaCod,6,3) = '" & Producto.gColConsuPrestAdm & "' THEN '20' WHEN SubString(ca.cCtaCod,6,3) = '" & Producto.gColHipoMiVivienda & "' THEN '23' Else '06' END + " _
                    & "   '0'+ SubString(cc.cLineaCred,6,1) + " _
                    & "   CASE WHEN SubString(ca.cCtaCod,6,3) IN ('" & Producto.gColConsuPrendario & "','" & Producto.gColConsuPrestAdm & "','" & Producto.gColHipoMiVivienda & "') THEN '01' ELSE SubString(ca.cCtaCod,7,2) END + " _
                    & "  '0'+SubString(cc.cLineaCred,2,1)+  CASE WHEN SUBSTRING(ca.cCtaCod,4,2) in ('99') THEN '11' ELSE SUBSTRING(ca.cCtaCod,4,2) END "
              
              
              'dbConec.AbreConexionRemota "07", False, False, "03"
            End If
        ElseIf pnBitCentral = False Then
            If sCodEmp = "112" Then
               'psSql = " SELECT '" & psClase & "'+ '0'+SubString(ca.cCodCta,3,1) + CASE WHEN SubString(ca.cCodCta,3,3) = '305' THEN '13' WHEN SubString(ca.cCodCta,3,3) = '320' THEN '20' WHEN SubString(ca.cCodCta,3,3) = '423' THEN '23' Else '06' END +" _
                     & "  '0'+ SubString(cc.cCodLinCred,5,1) + " _
                     & "  CASE WHEN SubString(ca.cCodCta,3,3) IN ('305','320','423') THEN '01' ELSE SubString(ca.cCodCta,4,2) END + " _
                     & "  '0'+SubString(cc.cCodLinCred,6,1) +  CASE WHEN LEFT(ca.cCodCta,2) in ('99') THEN '11' ELSE LEFT(ca.cCodCta,2) END Cta1, " _
                     & "  sum(cc.nIntDev) nSaldo, NULL Cta2, 0 nCtaSaldoImporte " _
                     & "  FROM CreditoAudi ca JOIN CreditoConsol cc ON cc.cCodCta = ca.cCodCta " _
                     & "  WHERE ca.cCodCta LIKE '_____" & pnMoneda & "%' and ca.cRefinan = 'N' and ca.cEstado in ('F','1','4','6','7') and cc.cEstado in ('F','1','4','6','7') " _
                     & "     and ( (SubString(ca.cCodCta,3,1) = 1 and ca.nDiasAtraso <= 15) or " _
                     & "      (SubString(ca.cCodCta,3,1) in (2,3,4) and ca.nDiasAtraso <= 30)  ) " _
                     & "  and ca.cCalGen in (3,4) " _
                     & "  group By '" & psClase & "'+  '0'+SubString(ca.cCodCta,3,1) + CASE WHEN SubString(ca.cCodCta,3,3) = '305' THEN '13' WHEN SubString(ca.cCodCta,3,3) = '320' THEN '20' WHEN SubString(ca.cCodCta,3,3) = '423' THEN '23' Else '06' END + " _
                     & "   '0'+ SubString(cc.cCodLinCred,5,1) + " _
                     & "   CASE WHEN SubString(ca.cCodCta,3,3) IN ('305','320','423') THEN '01' ELSE SubString(ca.cCodCta,4,2) END + " _
                     & "  '0'+SubString(cc.cCodLinCred,6,1)+  CASE WHEN LEFT(ca.cCodCta,2) in ('99') THEN '11' ELSE LEFT(ca.cCodCta,2) END "
               dbConec.AbreConexion 'Remota "07", False, False, "03"
             End If
        End If
   End If
   
   If pnTipoAsiento = 8 Then  'Provisiones de Carta Fianza
      If pnBitCentral = True Then

      ElseIf pnBitCentral = False Then
            If sCodEmp = "112" Or sCodEmp = "108" Or sCodEmp = "109" Then 'TRUJILLO / ICA
              psSql = "SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
                 & "FROM ( select '" & psClase & "'+ " _
                 & "          CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END + " _
                 & "          '0'+SubString(cCodCta,3,1) + " _
                 & "          CASE WHEN Left(cCodCta,2) in ('99') THEN '11' ELSE Left(cCodCta,2) END cCtaContCod , " _
                 & "          SUM( nProvision ) nSaldo " _
                 & "    FROM " & sServidorConsolidada & "CreditoAudi " _
                 & "    WHERE SubString(cCodCta,6,1) = '" & pnMoneda & "' and cEstado IN ('F','1','4','6','7','V') and SubString(cCodCta,4,2) in ('21') " _
                 & "    Group By '" & psClase & "'+ " _
                 & "       CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END + '0'+SubString(cCodCta,3,1) + " _
                 & "       CASE WHEN Left(cCodCta,2) in ('99') THEN '11' ELSE Left(cCodCta,2) END " _
                 & "    ) Cred FULL OUTER JOIN " _
                 & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) nCtaSaldoImporte  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
                 & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' " _
                 & "   ) ) cs on cs.cCtaContCod = Cred.cCtaContCod " _
                 & "ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
            End If
'            If scodemp = "102" Then 'LIMA
'              psSql = "SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
'                 & "FROM ( Select '" & psClase & "'+ " _
'                 & "          '0'+SubString(cCodCta,3,1) + " _
'                 & "          CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END cCtaContCod , " _
'                 & "          SUM( nProvision ) nSaldo " _
'                 & "    FROM " & sservidorconsolidada & "CreditoAudi " _
'                 & "    WHERE SubString(cCodCta,6,1) = '" & pnMoneda & "' and cEstado IN ('F','1','4','6','7','V') " _
'                 & "    Group By '" & psClase & "'+'0'+SubString(cCodCta,3,1) + " _
'                 & "       CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END + " _
'                 & "       CASE WHEN SubString(cCodCta,3,3)='305' THEN '13' WHEN SubString(cCodCta,3,3) = '320' THEN '20' WHEN SubString(cCodCta,3,3) = '423' THEN '23' Else '06' END + " _
'                 & "       CASE WHEN SubString(cCodCta,3,3) IN ('305','320','423') THEN '01' ELSE SubString(cCodCta,4,2) END + " _
'                 & "       '0'+SubString(cCodLinCred,6,1) + " _
'                 & "       CASE WHEN Left(cCodCta,2) in ('99') THEN '11' ELSE Left(cCodCta,2) END " _
'                 & "    ) Cred FULL OUTER JOIN " _
'                 & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) * -1 nCtaSaldoImporte  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
'                 & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' " _
'                 & "   ) ) cs on cs.cCtaContCod = Cred.cCtaContCod " _
'                 & "ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
'            End If
        End If
   End If
   
   If pnTipoAsiento = 9 Then  'Calificacion de Créditos Contigentes
        If pnBitCentral = True Then
        
        Else
            psSql = "SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
                  & "FROM ( " _
                  & "    SELECT '" & psClase & "010'+ convert(varchar(1), cCalGen + 1) cCtaContCod , " _
                  & "          SUM( nSaldoCap ) nSaldo " _
                  & "    From " & sServidorConsolidada & "CreditoAudi " _
                  & "    WHERE subString(cCodCta,6,1) = '" & pnMoneda & "' and cEstado IN ('F','1','4','6','7','V') " _
                  & "    group By '" & psClase & "010'+ convert(varchar(1), cCalGen + 1) " _
                  & "    UNION ALL " _
                  & "    SELECT '" & Left(psClase, 4) & "0305' cCtaContCod, SUM( nProvision ) nSaldo " _
                  & "    FROM " & sServidorConsolidada & "CreditoAudi " _
                  & "    WHERE subString(cCodCta,6,1) = '" & pnMoneda & "' and cEstado IN ('F','1','4','6','7','V') " _
                  & "     and cCalGen <> 0 " _
                  & "    ) Cred FULL OUTER JOIN " _
                  & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) nCtaSaldoImporte " _
                  & "      FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psClase & "%' and cs.dCtaSaldoFecha = " _
                  & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' ) " _
                  & "    ) cs on cs.cCtaContCod = Cred.cCtaContCod " _
                  & "ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
        End If
   End If
   
   If pnTipoAsiento = 10 Then  'Creditos en Garantia de Financiamiento
        If pnBitCentral = True Then
        
        Else
            psSql = "SELECT cred.cCtaContcod cta1, ISNULL(cred.nSaldo,0) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
                & "FROM ( select '" & psClase & "01' + '0'+left(cProdCod,1) cCtaContCod, " _
                & "         SUM( nSaldo ) nSaldo " _
                & "  From " & sServidorConsolidada & "CredSaldosCont " _
                & "  WHERE dFecha = '" & psFecha & "' and cConcepto = 'C' and cMoneda = '" & pnMoneda & "' and cFuenteFinanc IN ('3','5','6','7','8','9','A') " _
                & "  group By '" & psClase & "01' + '0'+left(cProdCod,1) " _
                & "  Union All " _
                & "  select '" & psClase & "02' + '0'+SubString(cCodCta,3,1) cCtaContCod , " _
                & "         SUM( nProvision * -1 ) nSaldo " _
                & "  From " & sServidorConsolidada & "CreditoAudi " _
                & "  WHERE SubString(cCodCta,6,1) = '" & pnMoneda & "' and cEstado IN ('F','1','4','6','7','V') and SubString(cCodLinCred,6,1) IN ('3','5','6','7','8','9','A') " _
                & "  group By '" & psClase & "02' + '0'+SubString(cCodCta,3,1) " _
                & " ) Cred FULL OUTER JOIN ( " _
                & "  SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & " , 2) nCtaSaldoImporte " _
                & "  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psClase & "%' and cs.dCtaSaldoFecha = " _
                & "     (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' ) " _
                & "     ) cs on cs.cCtaContCod = Cred.cCtaContCod " _
                & "Where IsNull(nSaldo, 0) <> 0 Or IsNull(nCtaSaldoImporte, 0) <> 0 " _
                & "ORDER BY cred.cCtaContCod"
        End If
   End If
   
   If pnTipoAsiento = 11 Then  'Provisiones de Carta Fianza
        If pnBitCentral = True Then
            If sCodEmp = "112" Or sCodEmp = "108" Or sCodEmp = "109" Then 'TRUJILLO /ICA
                '** Final
'    'ALPA 20090331*************************
'    psSql = " select y.cta1, y.nSaldo,y.cta2,y.nCtaSaldoImporte,y.nSaldo1,y.nSaldo2,y.ctaNuevaN,y.ctaNuevaP,"
'    psSql = psSql & "   y.nSaldoCtaProN,y.nSaldoCtaProProc,"
'    psSql = psSql & "       Case substring(Y.ctaNuevaN, 1, 1) "
'    psSql = psSql & "            when '6' then (y.nCtaSaldoImporte-y.nSaldo)"
'    psSql = psSql & "           else (y.nSaldoCtaProN-y.nSaldo1)"
'    psSql = psSql & "       End nSaldoProvision,"
'    psSql = psSql & "       Case substring(Y.ctaNuevaP, 1, 1)"
'    psSql = psSql & "             when '6' then (y.nCtaSaldoImporte-y.nSaldo)"
'    psSql = psSql & "             else (y.nSaldoCtaProProc-y.nSaldo2)"
'    psSql = psSql & "       End  nSaldoProvisionPro        "
'    psSql = psSql & "  From "
'    psSql = psSql & "   (   select  x.cta1,x.nSaldo,  x.cta2,x.nCtaSaldoImporte,"
'    psSql = psSql & "               x.nSaldo1,x.nSaldo2,x.ctaNuevaN,x.ctaNuevaP,  "
'    psSql = psSql & "               isnull((SELECT ROUND( nCtaSaldoImporte,2)   "
'    psSql = psSql & "               FROM ctasaldo cs"
'    psSql = psSql & "               WHERE cs.cCtaContCod = x.ctaNuevaN and"
'    psSql = psSql & "               cs.dCtaSaldoFecha = (   SELECT max(dCtaSaldoFecha)"
'    psSql = psSql & "                                       FROM CtaSaldo cs1 "
'    psSql = psSql & "                                       Where cs1.cCtaContCod = cs.cCtaContCod "
'    psSql = psSql & "                                       and cs1.dCtaSaldoFecha <= '" & psFecha & "'"
'    psSql = psSql & "                                   )"
'    psSql = psSql & "                       ),0) nSaldoCtaProN64,"
'    psSql = psSql & "  isnull((select top 1 round(isnull(be.nDebe,0),2)"
'    psSql = psSql & "  from balanceEstad be"
'    psSql = psSql & "  where be.cBalanceAnio = '" & DevuelveMesAnio(psFecha, 2) & "' "
'    psSql = psSql & "  and be.cBalanceMes= '" & DevuelveMesAnio(psFecha, 1) & "'"
'    psSql = psSql & "  and be.cCtaContCod=x.ctaNuevaN "
'    psSql = psSql & "  and be.cBalanceCate=1 "
'    psSql = psSql & "  and be.cBalanceTipo=1),0) nSaldoCtaProN,"
'    psSql = psSql & "    isnull((SELECT ROUND(isnull(nCtaSaldoImporte,0) ,2)"
'    psSql = psSql & "            FROM ctasaldo cs"
'    psSql = psSql & "           WHERE cs.cCtaContCod = x.ctaNuevaP and"
'    psSql = psSql & "                   cs.dCtaSaldoFecha = (   SELECT max(dCtaSaldoFecha)"
'    psSql = psSql & "                                           FROM CtaSaldo cs1"
'    psSql = psSql & "                                           Where cs1.cCtaContCod = cs.cCtaContCod"
'    psSql = psSql & "                                          and cs1.dCtaSaldoFecha <= '" & psFecha & "'"
'    psSql = psSql & "                                       )"
'    psSql = psSql & "            ),0) nSaldoCtaProProc64, "
'    psSql = psSql & "  isnull((select top 1 round(isnull(be.nDebe,0),2)"
'    psSql = psSql & "  from balanceEstad be"
'    psSql = psSql & "  where be.cBalanceAnio = '" & DevuelveMesAnio(psFecha, 2) & "' "
'    psSql = psSql & "  and be.cBalanceMes='" & DevuelveMesAnio(psFecha, 1) & "'"
'    psSql = psSql & "  and be.cCtaContCod=x.ctaNuevaP "
'    psSql = psSql & "  and be.cBalanceCate=1 "
'    psSql = psSql & "  and be.cBalanceTipo=1),0) "
'    psSql = psSql & "  nSaldoCtaProProc "
'    psSql = psSql & "       From  ("
'    psSql = psSql & "                 SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, "
'    psSql = psSql & "                   ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte,Round(ISNULL(cred.nSaldo1,0),3) nSaldo1,Round(ISNULL(cred.nSaldo2,0),3) nSaldo2, "
'    psSql = psSql & " case"
'    psSql = psSql & " when Round(IsNull(cred.nSaldo, 0), 3) >= IsNull(cs.nCtaSaldoImporte, 0)"
'    psSql = psSql & " then '43' + substring(isnull(cred.cCtaContcod,isnull(cs.cCtaContCod,'')), 3, 1) + '5' +  '02' + '01' + substring(isnull(cred.cCtaContcod,isnull(cs.cCtaContCod,'')), 9, 2) + '01' + right(isnull(cred.cCtaContcod,isnull(cs.cCtaContCod,'')), 2)"
'    psSql = psSql & "Else"
'    psSql = psSql & "'64' + substring(isnull(cred.cCtaContcod,isnull(cs.cCtaContCod,'')), 3, 1) + '1040201' + right(isnull(cred.cCtaContcod,isnull(cs.cCtaContCod,'')), 2)"
'    psSql = psSql & "End ctaNuevaN, "
'    psSql = psSql & "case "
'    psSql = psSql & " when Round(IsNull(cred.nSaldo, 0), 3) >= IsNull(cs.nCtaSaldoImporte, 0)"
'    psSql = psSql & " then '43' + substring(isnull(cred.cCtaContcod,isnull(cs.cCtaContCod,'')), 3, 1) + '5' +  '01' + '02' + substring(isnull(cred.cCtaContcod,isnull(cs.cCtaContCod,'')), 9, 2) + '01' + right(isnull(cred.cCtaContcod,isnull(cs.cCtaContCod,'')), 2)"
'    psSql = psSql & " Else"
'    psSql = psSql & " '64' + substring(isnull(cred.cCtaContcod,isnull(cs.cCtaContCod,'')), 3, 1) + '1040201'  + right(isnull(cred.cCtaContcod,isnull(cs.cCtaContCod,'')), 2)"
'    psSql = psSql & " End ctaNuevaP"
'    psSql = psSql & "                   FROM (  Select '" & psClase & "' + CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END + '09'+ Case substring(c.cCtaCod,6,3) When '121' Then '01' When '221' Then '02' End   + Substring(c.cCtaCod,4,2) cCtaContCod, "
'    'ALPA 20091006***************************************
''    psSql = psSql & "                   SUM(nProvision + nProvisionProciclica) nSaldo, "
'    psSql = psSql & "                   SUM(isnull(nProvision,0) + isnull(nProvisionProciclica,0)) nSaldo, "
''    psSql = psSql & "                   SUM( nProvision ) nSaldo1, "
'    psSql = psSql & "                   SUM( isnull(nProvision,0) ) nSaldo1, "
'    'psSql = psSql & "                   SUM(nProvisionProciclica) nSaldo2 "
'    psSql = psSql & "                   SUM(isnull(nProvisionProciclica,0) ) nSaldo2 "
'    '*****************************************************
'    psSql = psSql & "                   FROM ColocCalifProv c join "
'    psSql = psSql & "                   (Select distinct substring(E.cLineaCred,1,4) Lin, cCtacont from " & sServidorConsolidada & "ColocLineaCreditoEquiv E ) E "
'    psSql = psSql & "                       on substring(c.cLineaCred,1,4)  = lin "
'    psSql = psSql & "                   WHERE SubString(cCtaCod,9,1) = '" & pnMoneda & "' and nprdEstado in ('2020', '2021', '2022', '2030', '2031', '2032','2101', '2104', '2106', '2107', '2201','2205','2092')  And substring(cctacod,7,2) = '21' "
'    psSql = psSql & " Group By '" & psClase & "' + CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END + '09'+ Case substring(c.cCtaCod,6,3) When '121' Then '01' When '221' Then '02' End   + Substring(c.cCtaCod,4,2)"
'    '**************************************
'             '     psSql = " SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, " _
'              '        & " ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
'               '       & " FROM (Select '" & psClase & "' " _
'                '      & "    " _
'                '      & "   + CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END " _
'                '      & "    " _
'                '      & "    " _
'                '      & "   " _
'                '      & "   " _
'                '      & "   " _
'                '      & "   + Substring(c.cCtaCod,4,2) cCtaContCod , " _
'                '      & " SUM( nProvision ) nSaldo " _
'                '      & " FROM ColocCalifProv c join " _
'                '      & " (Select distinct substring(E.cLineaCred,1,4) Lin, cCtacont from " & sServidorConsolidada & "ColocLineaCreditoEquiv E ) E " _
'                '      & "   on substring(c.cLineaCred,1,4)  = lin " _
'                '      & "  WHERE SubString(cCtaCod,9,1) = '" & pnMoneda & "' and nprdEstado in ('2020', '2021', '2022', '2030', '2031', '2032','2101', '2104', '2106', '2107', '2201','2205')  And substring(cctacod,7,2) = '21' " _
'                '      & " Group By '" & psClase & "' " _
'                '      & "    " _
'                '      & "   + CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END " _
'                '      & "    " _
'                '      & "    " _
'                '      & "    " _
'                '      & "    " _
'                '      & "   " _
'                '      & " + Substring(cCtaCod,4,2) "
'
'
'                psSql = psSql & " "
'                psSql = psSql & " ) Cred FULL OUTER JOIN "
'                psSql = psSql & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) nCtaSaldoImporte  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "0[12]%' and cs.dCtaSaldoFecha = "
'                psSql = psSql & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & Format(psFecha, "YYYY/MM/DD") & "' "
'                psSql = psSql & "   ) ) cs on cs.cCtaContCod = Cred.cCtaContCod"
'                'ALPA 20090331*************************
'                'psSql = psSql & " ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
'                psSql = psSql & ") x"
'                psSql = psSql & ") y"
'                psSql = psSql & " ORDER BY ISNULL(cta2,cta1)"
'                '***************************************
            psSql = "exec B2_stp_sel_AjusteProvisionCartaFianza " & pnMoneda & ",'" & Format(psFecha, "YYYY/MM/YY") & "'," & pnTipCambio & ",'" & DevuelveMesAnio(psFecha, 1) & "','" & DevuelveMesAnio(psFecha, 2) & "'"
            
              End If
              If sCodEmp = "102" Then  'LIMA
                '** Final
                'psSql = "SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
                   & "FROM ( Select '" & psClase & "'+ " _
                   & "          '0'+SubString(cCtaCod,6,1) + " _
                   & "          CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END cCtaContCod , " _
                   & "          SUM( nProvision ) nSaldo " _
                   & "    FROM COlocCalifProv " _
                   & "    WHERE SubString(cCtaCod,9,1) = '" & pnMoneda & "' and nprdestado in ('" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "','" & gColPEstDesem & "', '" & gColPEstVenci & "', '" & gColPEstPRema & "', '" & gColPEstRenov & "', '" & gColocEstRecVigJud & "') " _
                   & "    Group By '" & psClase & "'+'0'+SubString(cCtaCod,6,1) + " _
                   & "          " _
                   & "          CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END" _
                   & "    ) Cred FULL OUTER JOIN " _
                   & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) * -1 nCtaSaldoImporte  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
                   & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' " _
                   & "   ) ) cs on cs.cCtaContCod = Cred.cCtaContCod " _
                   & "ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
              End If
        ElseIf pnBitCentral = False Then
            If sCodEmp = "112" Then  'TRUJILLO
              'psSql = "SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
                 & "FROM ( select '" & psClase & "'+ " _
                 & "          '0'+SubString(cCodCta,3,1) + " _
                 & "          CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END + " _
                 & "          CASE WHEN SubString(cCodCta,3,3)='305' THEN '13' WHEN SubString(cCodCta,3,3) = '320' THEN '20' WHEN SubString(cCodCta,3,3) = '423' THEN '23' Else '06' END + " _
                 & "          CASE WHEN SubString(cCodCta,3,3) IN ('305','320','423') THEN '01' ELSE SubString(cCodCta,4,2) END + " _
                 & "          CASE WHEN SubString(cCodLinCred,6,1) = 'A' THEN '10' WHEN SubString(cCodLinCred,6,1) = 'B' THEN '11' WHEN SubString(cCodLinCred,6,1) = 'C' THEN '12' ELSE '0'+SubString(cCodLinCred,6,1) END + " _
                 & "          CASE WHEN Left(cCodCta,2) in ('99') THEN '11' ELSE Left(cCodCta,2) END cCtaContCod , " _
                 & "          SUM( nProvision ) nSaldo " _
                 & "    FROM " & sServidorConsolidada & "CreditoAudi " _
                 & "    WHERE SubString(cCodCta,6,1) = '" & pnMoneda & "' and cEstado IN ('F','1','4','6','7','V') and not SubString(cCodCta,4,2) in ('21') " _
                 & "    Group By '" & psClase & "'+'0'+SubString(cCodCta,3,1) + " _
                 & "       CASE WHEN cCalGen = 0 THEN '02' ELSE '01' END + " _
                 & "       CASE WHEN SubString(cCodCta,3,3)='305' THEN '13' WHEN SubString(cCodCta,3,3) = '320' THEN '20' WHEN SubString(cCodCta,3,3) = '423' THEN '23' Else '06' END + " _
                 & "       CASE WHEN SubString(cCodCta,3,3) IN ('305','320','423') THEN '01' ELSE SubString(cCodCta,4,2) END + " _
                 & "       CASE WHEN SubString(cCodLinCred,6,1) = 'A' THEN '10' WHEN SubString(cCodLinCred,6,1) = 'B' THEN '11' WHEN SubString(cCodLinCred,6,1) = 'C' THEN '12' ELSE '0'+SubString(cCodLinCred,6,1) END + " _
                 & "       CASE WHEN Left(cCodCta,2) in ('99') THEN '11' ELSE Left(cCodCta,2) END " _
                 & "    ) Cred FULL OUTER JOIN " _
                 & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) * -1 nCtaSaldoImporte  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
                 & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' " _
                 & "   ) ) cs on cs.cCtaContCod = Cred.cCtaContCod " _
                 & "ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
            End If
        End If
   End If
    If pnTipoAsiento = 12 Then  'ajuste por calificacion de cartera
        If pnBitCentral = True Then
'            psSql = "       SELECT  BASE.cCtaContCod as cta1, "
'            psSql = psSql & "       IsNull(NSALDOCAP,0) as nSaldo,"
'            psSql = psSql & "       SALANT.CCTACONTCOD as cta2,"
'            psSql = psSql & "       IsNull(nCtaSaldoImporte,0) As nCtaSaldoImporte"
'            psSql = psSql & "   From "
'            psSql = psSql & "       (   SELECT CCTACONTCOD, CCALGEN, CMONEDA, SUM(NSALDOCAP) AS NSALDOCAP"
'            psSql = psSql & "           From"
'            psSql = psSql & "               (SELECT '81' + SUBSTRING(CCTACOD,9,1) + '601' +  CASE   WHEN CCALGEN=0  THEN '01'"
'            psSql = psSql & "                       WHEN CCALGEN=1  THEN '02'"
'            psSql = psSql & "                       WHEN CCALGEN=2  THEN '03'"
'            psSql = psSql & "                       WHEN CCALGEN=3  THEN '04'"
'            psSql = psSql & "                       WHEN CCALGEN=4  THEN '05' END + '0' + SUBSTRING(CCTACOD,6,1) + SUBSTRING(CCTACOD,4,2) AS CCTACONTCOD,"
'            psSql = psSql & "                       CCALGEN, SUBSTRING(CCTACOD,9,1) AS CMONEDA,"
'            psSql = psSql & "                       NSALDOCAP"
'            psSql = psSql & "                From COLOCCALIFPROV"
'            psSql = psSql & "            WHERE   NPRDESTADO IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205) And SUBSTRING(CCTACOD,9,1) = '" & pnMoneda & "' And  SUBSTRING(CCTACOD,6,3) Not In ('121','221')  "
'            psSql = psSql & "   ) AS DAT"
'            psSql = psSql & "            GROUP BY CCTACONTCOD, CCALGEN, CMONEDA ) AS BASE"
'            psSql = psSql & "   Full Outer Join"
'            psSql = psSql & "            (SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) nCtaSaldoImporte "
'            psSql = psSql & "             FROM    CTASALDO C"
'            psSql = psSql & "             WHERE   cCtaContCod  LIKE '" & psCtaDebe & "%'"
'            psSql = psSql & "                     AND  dCtaSaldoFecha   = (SELECT MAX(dCtaSaldoFecha)"
'            psSql = psSql & "                                              FROM  CTASALDO C1"
'            psSql = psSql & "                       WHERE   C1.cCtaContCod = C.cCtaContCod and dCtaSaldoFecha<= '" & psFecha & "' )) AS SALANT"
'            psSql = psSql & "                               ON SALANT.cCtaContCod = BASE.CCTACONTCOD "
            psSql = "exec B2_stp_sel_AjusteCalificacionDeCreditos 81," & pnMoneda & ",'" & Format(psFecha, "YYYY/MM/YY") & "'," & pnTipCambio
        End If
    End If
   
    If pnTipoAsiento = 13 Then  'ajuste por calificacion de cartera
        If pnBitCentral = True Then
            psSql = "      SELECT   BASE.CCTACONT AS CTA1, NMOV AS NSALDO, "
            psSql = psSql & "       CCTACONTCOD AS CTA2, nCtaSaldoImporte,"
            psSql = psSql & "       CASE WHEN SUBSTRING(CCTACONTCOD,3,1) ='1' THEN nCtaSaldoImporte ELSE nCtaSaldoImporteME END AS NSALANT,"
            psSql = psSql & "       NMOV - CASE WHEN SUBSTRING(CCTACONTCOD,3,1) ='1' THEN nCtaSaldoImporte ELSE nCtaSaldoImporteME END AS NDMOV"
            psSql = psSql & " From"
            psSql = psSql & "           (   SELECT CCTACONT, CMONEDA, -1 * SUM(nProvAutoCero + nProvMVCero + nProvMV50 +nGarCTS + NPROV100) AS NMOV"
            psSql = psSql & "               From"
            psSql = psSql & "                   ("
            psSql = psSql & "                       SELECT  '81' + CMONEDA + '603' + CASE   WHEN NTIPO =0 THEN '01'"
            psSql = psSql & "                                                               WHEN NTIPO =50 THEN '04'"
            psSql = psSql & "                                                               WHEN NTIPO = 100 THEN '05' END AS CCTACONT,*"
            psSql = psSql & "                       From"
            psSql = psSql & "                           ("
            psSql = psSql & "                               SELECT  SUBSTRING(CCTACOD,9,1) AS CMONEDA, NPROVISION,"
            psSql = psSql & "                                       CASE    WHEN nProvAutoCero >0 THEN 0"
            psSql = psSql & "                                               WHEN nProvMVCero >0 THEN 0"
            psSql = psSql & "                                               WHEN nProvMV50>0 THEN 50"
            psSql = psSql & "                                               WHEN nGarCTS>0 THEN 50"
            psSql = psSql & "                                       ELSE 100 END NTIPO,"
            psSql = psSql & "                                       nProvAutoCero,"
            psSql = psSql & "                                       nProvMVCero,"
            psSql = psSql & "                                       nProvMV50,"
            psSql = psSql & "                                       nGarCTS,"
            psSql = psSql & "                                       CASE WHEN nProvAutoCero + nProvMVCero +nProvMV50 + nGarCTS=0 THEN  NPROVISION ELSE 0 END NPROV100"
            psSql = psSql & "                                From"
            psSql = psSql & "                                   (SELECT  *,"
            psSql = psSql & "                                           CASE    WHEN nGarant = 1 and (SELECT COUNT(*)"
            psSql = psSql & "                                           FROM    " & sServidorConsolidada & "GarantiasConsol G"
            psSql = psSql & "                                                   JOIN " & sServidorConsolidada & "GarantCredConsol GC ON G.cNumGarant = GC.cNumGarant"
            psSql = psSql & "                                                   where  G.nTipoGarant = 6 and G.cDocGarant = 121 and GC.cCtaCod = C.cCtaCod )= 0 THEN nProvision ELSE 0 END AS nProvAutoCero,"
            psSql = psSql & "                                           0 AS nProvMVCero,"
            psSql = psSql & "                                           CASE    WHEN SUBSTRING(CCTACOD,6,3) ='423' THEN ROUND(nProvision,0) ELSE 0 END AS nProvMV50,"
            psSql = psSql & "                                           CASE    WHEN nGarant=1 AND (    SELECT COUNT(*)"
            psSql = psSql & "                                                                           FROM   " & sServidorConsolidada & "GarantiasConsol G"
            psSql = psSql & "                                                                           JOIN " & sServidorConsolidada & "GarantCredConsol GC ON G.cNumGarant = GC.cNumGarant"
            psSql = psSql & "                                                                           where  G.nTipoGarant = 6 and G.cDocGarant = 121 and GC.cCtaCod = C.cCtaCod )> 0 THEN nProvision ELSE 0 END NGARCTS"
            psSql = psSql & "                                     FROM    COLOCCALIFPROV  C"
            psSql = psSql & "                                     WHERE   nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107)"
            psSql = psSql & "                                     Union All"
            psSql = psSql & "                                     select *, 0,0,0,0 from coloccalifprov where cCalGen = '0'  and nrcambiario= 0 and substring(cctacod,9,1) = '" & pnMoneda & "'"
            psSql = psSql & "                                       )AS X"
            psSql = psSql & "                               )AS Y"
            psSql = psSql & "                           )AS Z"
            psSql = psSql & "           GROUP BY CCTACONT, CMONEDA ) AS BASE"
            psSql = psSql & "       JOIN"
            psSql = psSql & "       (SELECT     cCtaContCod,ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) AS nCtaSaldoImporte, nCtaSaldoImporteME"
            psSql = psSql & "        FROM       CTASALDO C"
            psSql = psSql & "        WHERE      cCtaContCod  LIKE '" & psCtaDebe & "%'"
            psSql = psSql & "                   AND  dCtaSaldoFecha   = (   SELECT MAX(dCtaSaldoFecha)"
            psSql = psSql & "                                               FROM  CTASALDO C1"
            psSql = psSql & "                                               WHERE  C1.cCtaContCod = C.cCtaContCod and dCtaSaldoFecha<= '" & psFecha & "')) AS SALANT"
            psSql = psSql & "         ON SALANT.CCTACONTCOD  = BASE.CCTACONT"
            psSql = psSql & "       ORDER BY CMONEDA, CCTACONT"
        End If
    End If
   
    If pnTipoAsiento = 14 Then
        If pnBitCentral = True Then
            
'                psSql = " SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.Saldo,0),3) nSaldo, cs.cCtaContCod cta2, " _
'                      & " ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
'                      & " FROM (select Sum(nSaldCnt) Saldo, " _
'                      & " case substring(a.cctacod,6,3) when '232' then case isnull(binactiva,0) when 1 then  '21" & Trim(Str(pnMoneda)) & "202' else '21" & Trim(Str(pnMoneda)) & "201' end when '233' then '21" & Trim(Str(pnMoneda)) & "303' when '234' then '21" & Trim(Str(pnMoneda)) & "305' end" _
'                      & " + case when npersoneria in (1,2) and substring(a.cctacod,6,3) <> '234' then '01' else '' end + Case when npersoneria in (2,3) then '02'" _
'                      & "  else Case when substring(a.cctacod,6,3) <> '234' then '01' else case isnull(cCodInst,'') when '1090100012521' then '02' else '01' end  end end + substring(a.cctacod,4,2) cCtaContCod" _
'                      & "  from capsaldosdiarios a" _
'                      & "  inner join captaciones b on a.cctacod = b.cctacod" _
'                      & "  left join captaccts c on a.cctacod = c.cctacod" _
'                      & "  where dfecha >= '" & psFecha & "' and dfecha < DateAdd(Day,1,'" & psFecha & "') and a.cctacod like '________" & Trim(Str(pnMoneda)) & "%' and a.cCtaCod Not In (SELECT PC.CCTACOD FROM DBConsolidada..ProductoBloqueosConsol PC WHERE cMovNroDbl IS NULL AND nBlqMotivo = 3 And PC.cCtaCod Like '_____233%' And dCierre = '" & psFecha & "') " _
'                      & "  and npersoneria in (1,2,3) group by" _
'                      & "  case substring(a.cctacod,6,3) when '232' then case isnull(binactiva,0) when 1 then  '21" & Trim(Str(pnMoneda)) & "202' else '21" & Trim(Str(pnMoneda)) & "201' end when '233' then '21" & Trim(Str(pnMoneda)) & "303' when '234' then '21" & Trim(Str(pnMoneda)) & "305' end" _
'                      & "  + case when npersoneria in (1,2) and substring(a.cctacod,6,3) <> '234' then '01' else '' end + Case when npersoneria in (2,3) then '02'" _
'                      & "         else Case when substring(a.cctacod,6,3) <> '234' then '01' else case isnull(cCodInst,'') when '1090100012521' then '02' else '01' end  end end + substring(a.cctacod,4,2)"
'
'                 'psSql = psSql & "Union All select Sum(nSaldCnt) Saldo, " _
'                      & " case substring(a.cctacod,6,3) when '232' then case '23" & Trim(Str(pnMoneda)) & "201' when '233' then '21" & Trim(Str(pnMoneda)) & "303' when '234' then '21" & Trim(Str(pnMoneda)) & "305' end" _
'                      & " + Case npersoneria when 3 then '02'" _
'                      & "  else Case when substring(a.cctacod,6,3) <> '234' then '01' else case isnull(cCodInst,'') when '1060300008399' then '02' else '01' end  end end + substring(a.cctacod,4,2) cCtaContCod" _
'                      & "  from capsaldosdiarios a" _
'                      & "  inner join captaciones b on a.cctacod = b.cctacod" _
'                      & "  left join captaccts c on a.cctacod = c.cctacod" _
'                      & "  where dfecha >= '" & psFecha & "' and dfecha < DateAdd(Day,1,'" & psFecha & "') and a.cctacod like '________" & Trim(Str(pnMoneda)) & "%' and a.cCtaCod Not In (Select cCtaCod from dbconsolidada..CapSaldosCtasBlq where dfecha = '" & psFecha & "' )" _
'                      & "  and npersoneria in (4,5,6) group by" _
'                      & "  case substring(a.cctacod,6,3) when '232' then case isnull(binactiva,0) when 1 then  '21" & Trim(Str(pnMoneda)) & "202' else '21" & Trim(Str(pnMoneda)) & "201' end when '233' then '21" & Trim(Str(pnMoneda)) & "303' when '234' then '21" & Trim(Str(pnMoneda)) & "305' end" _
'                      & "  + Case npersoneria when 3 then '02'" _
'                      & "         else Case when substring(a.cctacod,6,3) <> '234' then '01' else case isnull(cCodInst,'') when '1060300008399' then '02' else '01' end  end end + substring(a.cctacod,4,2)"
'
'                      psSql = psSql & "  Union All" _
'                      & "  select Sum(nSaldCnt) Saldo, '21" & Trim(Str(pnMoneda)) & "70401' + Case npersoneria when 3 then '01' else '01' end  + substring(a.cctacod,4,2) cCtaContCod" _
'                      & "  from capsaldosdiarios a" _
'                      & "  inner join captaciones b on a.cctacod = b.cctacod left join captaccts c on a.cctacod = c.cctacod" _
'                      & "  where dfecha >= '" & psFecha & "' and dfecha < DateAdd(Day,1,'" & psFecha & "') and a.cctacod like '_____23[3]" & Trim(Str(pnMoneda)) & "%' and a.cCtaCod In (SELECT PC.CCTACOD FROM DBConsolidada..ProductoBloqueosConsol PC WHERE cMovNroDbl IS NULL AND nBlqMotivo = 3 And PC.cCtaCod Like '_____233%' and dCierre = '" & psFecha & "') and npersoneria in (1,2,3) group by" _
'                      & "  '21" & Trim(Str(pnMoneda)) & "70401' + Case npersoneria when 3 then '01' else '01' end  + substring(a.cctacod,4,2)"
'                psSql = psSql & " " _
'                      & " ) Cred FULL OUTER JOIN " _
'                      & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) nCtaSaldoImporte  FROM ctasaldo cs WHERE (cs.cctacontcod like '21" & Trim(Str(pnMoneda)) & "[237]%') and cs.dCtaSaldoFecha = " _
'                      & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' " _
'                      & "   ) ) cs on cs.cCtaContCod = Cred.cCtaContCod" _
'                      & " ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"

                '*** PEAC 20130111
                psSql = "exec stp_sel_AjusteSaldosAhorro '" & Trim(Str(pnMoneda)) & "','" & psFecha & "'," & pnTipCambio
      
        End If
    End If
   
    If pnTipoAsiento = 15 Then
        If pnBitCentral = True Then
                'psSql = " sELECT a.cCtaCod cta1, b.nProvisionSoles nSaldo, a.cCtaCod cta2, b.nProvisionSoles nCtaSaldoImporte FROM dbConsolidada..CreditoConsolTotal a" _
                      & " Inner Join (Select cCtaCod, nProvisionSoles,  from dbConsolidada..ColocCalifProvTotal where dfecha = '" & Year(psFecha) & "/12/31') b on a.cCtaCod = b.cCtaCod" _
                      & " WHERE datediff(month,dCancelado,'" & psFecha & "') = 0"

                psSql = " Select cCtaContCod14 cta1, cCtaContCod64 cta2, Sum(Monto) nSaldo, 0 nCtaSaldoImporte From" _
                      & " (SELECT  '14" & Trim(Str(pnMoneda)) & "9' + '0' + Substring(a.cCtaCod,6,1) + '0106' + '01' + (SELECT LTRIM(RTRIM(cEquivAsiento)) From dbcmac..AsientoProdEquiv AE Where cEquivTpo ='SC' and AE.cProducto = Substring(a.cCtaCod,6,3)) + cCtacont + Substring(a.cCtaCod,4,2) cCtaContCod14," _
                      & "          '64" & Trim(Str(pnMoneda)) & "1040201' + '0' + Substring(a.cCtaCod,6,1) + (SELECT LTRIM(RTRIM(cEquivAsiento)) From dbcmac..AsientoProdEquiv AE Where cEquivTpo ='SC' and AE.cProducto = Substring(a.cCtaCod,6,3)) + cCtacont + Substring(a.cCtaCod,4,2) cCtaContCod64," _
                      & "           b.nProvisionSoles Monto FROM dbconsolidada..CreditoConsolTotal a" _
                      & "  Inner Join (Select cCtaCod, nProvision nProvisionSoles, cLineaCred from dbconsolidada..ColocCalifProvTotal where dfecha = '" & Trim(Str(CInt(Year(psFecha) - 1))) & "/12/31' And cCalGen <> 0  And cCtaCod Like '________" & Trim(Str(pnMoneda)) & "%' ) b on a.cCtaCod = b.cCtaCod" _
                      & "  Inner Join (Select distinct substring(E.cLineaCred,1,4) Lin, cCtacont from dbconsolidada..ColocLineaCreditoEquiv E) aaa on lin = left(b.cLineaCred,4)" _
                      & "  WHERE datediff(month,dCancelado,'" & psFecha & "') = 0) As AA" _
                      & "  Group by cCtaContCod14, cCtaContCod64"
        End If
    End If
   
   If pnTipoAsiento = 16 Then  'Comision de Creditos Castigados
        If pnBitCentral = True Then
             If sCodEmp = "112" Or sCodEmp = "108" Or sCodEmp = "109" Then 'TRUJILLO / ICA
                  psSql = "SELECT cred.cCtaContcod cta1, Round(ISNULL(cred.nSaldo,0),3) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
                     & " FROM ( select '" & psClase & "' " _
                     & "     + '0'+SubString(c.cCtaCod,6,1)  " _
                     & "     + CASE WHEN SubString(c.cCtaCod,6,3) = '" & Producto.gColHipoMiVivienda & "' THEN '23' Else '06' END " _
                     & "     + '0' + SubString(c.cLineaCred,6,1)  + cEquivAsiento +  left(E.cCtaCont,2) " _
                     & "     + Substring(c.cCtaCod,4,2) cCtaContCod, " _
                     & "     Sum(nGastoPend) nSaldo " _
                     & "    FROM " & sServidorConsolidada & "CreditoConsol c Join " _
                     & "    (Select distinct substring(E.cLineaCred,1,4) Lin, cCtacont from " & sServidorConsolidada & "ColocLineaCreditoEquiv E ) E " _
                     & "    On substring(c.cLineaCred,1,4)  = lin Join (SELECT LTRIM(RTRIM(cEquivAsiento)) cEquivAsiento, cProducto From AsientoProdEquiv AE Where cEquivTpo ='SC') AE On AE.cProducto = SubString(c.cCtaCod,6,3) " _
                     & "    Inner Join " & sServidorConsolidada & "CreditoSaldoConsol CS On dfecha = '" & psFecha & "' And CS.cCtaCod = c.cCtaCod " _
                     & "  WHERE substring(C.cCtaCod,9,1) = '" & pnMoneda & "' and C.nPrdEstado IN('" & gColocEstRecVigCast & "','2206') and C.nSaldoCap >= 0 " _
                     & "  Group By '" & psClase & "'+ '0'+SubString(c.cCtaCod,6,1) + " _
                     & "     + CASE WHEN SubString(c.cCtaCod,6,3) = '" & Producto.gColHipoMiVivienda & "' THEN '23' Else '06' END " _
                     & "     + '0' + SubString(c.cLineaCred,6,1) + cEquivAsiento +  left(E.cCtaCont,2)  " _
                     & "     + Substring(c.cCtaCod,4,2)  " _
                     & "    ) Cred FULL OUTER JOIN " _
                     & "    ( SELECT cCtaContCod, ROUND( nCtaSaldoImporte " & IIf(pnMoneda = 2, " / " & pnTipCambio, "") & ",2) nCtaSaldoImporte  FROM ctasaldo cs WHERE cs.cCtaContCod like '" & psCtaDebe & "%' and cs.dCtaSaldoFecha = " _
                     & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' " _
                     & "   ) ) cs on cs.cCtaContCod = Cred.cCtaContCod " _
                     & "ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod)"
             End If
        End If
    End If
   'ALPA 20090709************
   If pnTipoAsiento = 17 Then  'Asiento Final
        If pnBitCentral = True Then
            Dim dFechaAnte As Date
            dFechaAnte = DateAdd("m", 1, "01/" & Format(DevuelveMesAnio(psFecha, 1), "00") & "/" & Format(DevuelveMesAnio(psFecha, 2), "0000")) - 1
            psSql = "exec  stp_sel_AsientoDeCierre '" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "','" & Format(dFechaAnte, "YYYY/MM/DD") & "'," & pnTipCambio
        End If
    End If
   '**************************
      'ALPA 20090709************
   If pnTipoAsiento = 18 Then  'Asiento Final
        If pnBitCentral = True Then
            psSql = "exec  B2_stp_sel_AjusteInteresesDiferidos '" & Format(psFecha, "YYYY/MM/DD") & "','" & pnMoneda & "'," & pnTipCambio
        End If
    End If
    If pnTipoAsiento = 19 Then  'Asiento Final
        If pnBitCentral = True Then
            psSql = "exec  B2_stp_sel_AsientoGarantiasCartasFianzas '" & Format(psFecha, "YYYY/MM/DD") & "','" & pnMoneda & "'," & pnTipCambio
        End If
    End If

   '**************************
   'MADM 20110805
    If pnTipoAsiento = 20 Then  'Asiento Final
        If pnBitCentral = True Then
            psSql = "exec  B2_stp_sel_AsientoSaldoGarantiasCartasFianzas '" & Format(psFecha, "YYYY/MM/DD") & "','" & pnMoneda & "'," & pnTipCambio
        End If
    End If
'END MADM
    If pnTipoAsiento = 21 Then  'Asiento Final
        If pnBitCentral = True Then
            psSql = "exec  B2_stp_sel_AjusteInteresesDiferidosSaldoMigrado '" & Format(psFecha, "YYYY/MM/DD") & "','" & pnMoneda & "'," & pnTipCambio
        End If
    End If
    'EJVG20130307 ***
    If pnTipoAsiento = 23 Then  'Comisión CF
        If pnBitCentral = True Then
            psSql = "exec  stp_sel_AjusteComisionCF '" & Format(psFecha, "YYYY/MM/DD") & "','" & pnMoneda & "'," & pnTipCambio
        End If
    End If
    'END EJVG *******
    'ALPA20131228 ***
    If pnTipoAsiento = 24 Then  'INTERESES DIFERIDOS FINALES
        If pnBitCentral = True Then
            psSql = "exec  B2_stp_sel_AjusteInteresesDiferidosSaldoFinal '" & Format(psFecha, "YYYY/MM/DD") & "','" & pnMoneda & "'," & pnTipCambio
        End If
    End If
    'END ALPA *******
    'ALPA20141202 ***
    If pnTipoAsiento = 25 Then  'INTERESES DIFERIDOS FINALES
        If pnBitCentral = True Then
            psSql = "exec  stp_sel_GenerarAjusteCompraDeuda '" & Format(psFecha, "YYYY/MM/DD") & "','" & pnMoneda & "'," & pnTipCambio
        End If
    End If
    'END ALPA *******
    'ALPA20141204 ***
    If pnTipoAsiento = 26 Then  'RECLASIFICACION DE PROVISIONES PROCICLICAS
        If pnBitCentral = True Then
            psSql = "exec  stp_sel_ReclasificaciónProvisionesProciclicas '14','" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "'," & pnTipCambio & ""
        End If
    End If
    'ALPA20150813 ***
    If pnTipoAsiento = 27 Then  'RECLASIFICACION DE CREDITOS REFINANCIADOS
        If pnBitCentral = True Then
            psSql = "exec  B2_stp_sel_AjusteInteresesDiferidosSaldoFinal_CreditosReclasificados '" & Format(psFecha, "YYYY/MM/DD") & "','" & pnMoneda & "'," & pnTipCambio
        End If
    End If
    'PASI 20170412 ************************************
    If pnTipoAsiento = 28 Then
        If pnBitCentral = True Then
            psSql = "exec B2_stp_sel_ReversionInteresesDevengadosxZonaInundada '51', '" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "'," & pnTipCambio
        End If
    End If
    'PASI END******************************************
    '***************NAGL 202007 Según ACTA N°049-2020************
    If pnTipoAsiento = 29 Then
        If pnBitCentral = True Then
            Dim dFechaAnteFae As Date
            dFechaAnteFae = DateAdd("m", 1, "01/" & Format(DevuelveMesAnio(psFecha, 1), "00") & "/" & Format(DevuelveMesAnio(psFecha, 2), "0000")) - 1
            psSql = "exec stp_sel_AsientoCierreFAEMype '" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "','" & Format(dFechaAnteFae, "YYYY/MM/DD") & "'," & pnTipCambio
        End If
    End If
    If pnTipoAsiento = 30 Then
        If pnBitCentral = True Then
            Dim dFechaAnteReac As Date
            dFechaAnteReac = DateAdd("m", 1, "01/" & Format(DevuelveMesAnio(psFecha, 1), "00") & "/" & Format(DevuelveMesAnio(psFecha, 2), "0000")) - 1
            psSql = "exec stp_sel_AsientoCierreReactiva '" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "','" & Format(dFechaAnteReac, "YYYY/MM/DD") & "'," & pnTipCambio
        End If
    End If
    If pnTipoAsiento = 31 Then
        If pnBitCentral = True Then
            psSql = "exec stp_sel_AsientoCierreCarteraReprog '" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "'," & pnTipCambio
        End If
    End If
    '*************************************************************
    '***************NAGL 202102 Según ACTA N°017-2021************
    If pnTipoAsiento = 34 Then
        If pnBitCentral = True Then
            psSql = "exec stp_sel_AsientoCierreProvReprogCOVID '" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "', " & pnTipCambio & ", '" & psRep & "'"
        End If
    End If
    '************************************************************
   Set AjusteInteresesColocaciones = dbConec.CargaRecordSet(psSql)
   Set dbConec = Nothing
Exit Function
AjusteInteresesColocacionesErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:AjusteInteresesColocaciones Method")
End Function

Public Function CargaEstadGarantias(psFecha As String, pnMoneda As Integer, pnTipCambio As Currency, Optional ByVal pnTipo As Integer) As ADODB.Recordset 'By Capi 10122007 de acuerdo a nuevo parametro
Dim nCantGar As Integer
Dim psUltCod As String
Dim psSqlCab As String
Dim nRelacion As Integer

On Error GoTo CargaEstadGarantiasErr
   Set dbConec = New DConecta
   'By Capi 10122007 para el tratamiento de avales
   dbConec.AbreConexion
      
   
   
   
   If pnTipo = 1 Then
        nRelacion = 20
        If pnBitCentral = True Then
            psSql = "Select Distinct convert(int,cTipoGarant) cTipoGarant, replace(replace(replace(replace(replace(replace(c.cConsDescripcion,' ','_'),'.',''),'/','_'),')',''),'(',''),':','') cConsDescripcion from " & sServidorConsolidada & "CredGarantiasConsol cg JOIN Constante c ON c.nConsValor = convert(int,cg.cTipoGarant) where c.nConsCod = " & gPersGarantia & " and datediff(d,dFecha,'" & psFecha & "') = 0 Order by cTipoGarant "
            Set prs = dbConec.CargaRecordSet(psSql)
            If prs.EOF Then
                Err.Raise 50001, "CargaEstadGarantías", "No existen datos de Garantías"
            End If
            nCantGar = prs.RecordCount
            psSqlCab = "SELECT M.Prod, M.FF Fte_Finan, M.Ag Agencia"
            psSql = "FROM ( SELECT DISTINCT substring(cg.cCtaCod,6,3) Prod, CASE WHEN cFF is NULL THEN '01' ELSE cFF END FF, substring(cg.cCtaCod,4,2) Ag FROM " & sServidorConsolidada & "CredGarantiasConsol cg WHERE nMoneda = '" & pnMoneda & "' and datediff(d,dFecha,'" & psFecha & "') = 0 ) M "
            Do While Not prs.EOF
                psSqlCab = psSqlCab & ", " & prs!cConsDescripcion & " "
                psUltCod = "c_" & prs!cTipoGarant
            
                '& "    Sum( ROUND(  (nMonto / (  (Case isnUll(CC.nSaldoCap,1) when 0 then 1 else isnUll(CC.nSaldoCap,1) end )  / (Case isNull(D.SaldoCap,1) when 0 then 1 else isNull(D.SaldoCap,1) end  ) )  ) * " & IIf(pnMoneda = 1, 1, pnTipCambio) & ", 2)) " & prs!cConsDescripcion & " " _

                '& " Inner Join " & sServidorConsolidada & "GarantCredConsol GCC1 On GCC1.cCtaCod=CG.cCtaCod " _
                '& " Inner Join " & sServidorConsolidada & "GarantCredConsol GCC2 On GCC2.cPersCod=PP.cPersCod" _

                '**Modificado por DAOR 20080404 ********************************************************
                'psSql = psSql & " LEFT JOIN ( SELECT substring(cg.cCtaCod,6,3) Prod, " _
                '& "    CASE WHEN cFF is NULL THEN '01' ELSE cFF END FF, substring(cg.cCtaCod,4,2) Ag , Sum(D.MontoApr) Desembolso, " _
                '& "    Sum(ROUND(nMonto*ROUND( CASE WHEN SUBSTRING(CC.CCTACOD,9,1)='1' THEN CC.nSaldoCap ELSE ROUND(CC.nSaldoCap*" & pnTipCambio & ",2) END / (CASE WHEN ISNULL(D.SaldoCap,0)=0 THEN 1 ELSE ISNULL(D.SaldoCap,0) END), 2)*" & IIf(pnMoneda = 1, 1, pnTipCambio) & ",2)) " & prs!cConsDescripcion & " " _
                '& " FROM " & sServidorConsolidada & "CredGarantiasConsol cg " _
                '& " Inner Join " & sServidorConsolidada & "ProductoPersonaConsol PP On CG.cCtaCod = PP.cCtaCod" _
                '& " Inner Join " & sServidorConsolidada & "CreditoConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac =" & nRelacion & "" _
                '& " Inner Join (select  GC.cPersCod,Sum(  Case Substring(B.cctacod,9,1) When '1' then B.nMontoApr else B.nMontoApr * " & pnTipCambio & " end) MontoApr, Sum(  Case Substring(B.cctacod,9,1) When '1' then B.nSaldoCap else B.nSaldoCap * " & pnTipCambio & " end) SaldoCap, Count(*) Num" _
                ' & "             from    " & sServidorConsolidada & "ProductoPersonaConsol GC" _
                ' & "             INNER jOIN " & sServidorConsolidada & "creditoconsol B On GC.cCtaCod = B.cCtaCod and GC.nPrdPersRelac=" & nRelacion & "" _
                ' & "          group by GC.cPersCod) As D On PP.cPersCod = D.cPersCod" _
                ' & " WHERE cg.nMoneda = '" & pnMoneda & "' and datediff(d,CG.dFecha,'" & psFecha & "') = 0 and convert(int,cTipoGarant) = " & prs!cTipoGarant & " " _
                ' & "GROUP BY substring(cg.cCtaCod,6,3), CASE WHEN cFF is NULL THEN '01' ELSE cFF END, substring(cg.cCtaCod,4,2) " _
                ' & " ) " & psUltCod & " ON " & psUltCod & ".Prod = M.Prod and " & psUltCod & ".FF = M.FF and " & psUltCod & ".Ag = M.Ag "
                
                psSql = psSql & " left join ( select substring(cg.cCtaCod,6,3) Prod, " _
                & "     case when cFF is null then '01' else cFF end FF, substring(CG.cCtaCod,4,2) Ag , " _
                & "     Sum(round(CG.nMonto,2)) " & prs!cConsDescripcion & " " _
                & " from " & sServidorConsolidada & "CredGarantiasConsol CG inner join " & sServidorConsolidada & "ProductoPersonaConsol PP On CG.cCtaCod = PP.cCtaCod " _
                & "     inner join " & sServidorConsolidada & "CreditoConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac =" & nRelacion & "" _
                & " where CG.nMoneda = '" & pnMoneda & "' and datediff(d,CG.dFecha,'" & psFecha & "') = 0 and convert(int,CG.cTipoGarant) = " & prs!cTipoGarant & " " _
                & " group by substring(CG.cCtaCod,6,3), case when cFF is null then '01' else cFF end, substring(CG.cCtaCod,4,2) " _
                & " ) " & psUltCod & " ON " & psUltCod & ".Prod = M.Prod and " & psUltCod & ".FF = M.FF and " & psUltCod & ".Ag = M.Ag "
                '*******************************************************************************************************************************
                
                prs.MoveNext
            Loop
            psSql = psSqlCab & psSql & " ORDER BY M.Prod, M.FF, M.Ag"
        Else
            If Not dbConec.AbreConexion Then Exit Function 'Remota("07", False, False, "03")
                'ALPA20130720
                'Se adicionó el item nActivoLeasing para activos leasing
                psSql = "SELECT substring(cg.cCodCta,3,3) Prod, " _
                & "  CASE WHEN cFF is NULL THEN '01' ELSE cFF END Fte_Finan, substring(cg.cCodCta,1,2) Agencia , " _
                & "  SUM(nHipoteca) Hipoteca, SUM(nVehicular) Vehicular, SUM(nIndustrial) Industrial, SUM(nAgricola) Agricola, SUM(nCartaFianza) CartaFianza, SUM(nDeposito) Deposito, SUM(nJoyas) Joyas , SUM(nGlobalFlo) Global_Flotante, SUM(nFianzaSolida) Fianza_Solidar, SUM(nOtrosG) Otros, SUM(nPersonal) Personal, SUM(nAutorizacion) Autorizacion,SUM(nActivoLeasing) nActivoLeasing " _
                & "FROM Credgarantias cg " _
                & "WHERE cMoneda = '" & pnMoneda & "' and datediff(d,dFecha,'" & psFecha & "') = 0 " _
                & "GROUP BY substring(cg.cCodCta,3,3), CASE WHEN cFF is NULL THEN '01' ELSE cFF END, substring(cg.cCodCta,1,2) " _
                & "ORDER BY substring(cg.cCodCta,3,3), CASE WHEN cFF is NULL THEN '01' ELSE cFF END, substring(cg.cCodCta,1,2) "
        End If
    Else
        'JUEZ 20131107 ***************************************************************************************
        'nRelacion = 25
        'psSql = " SELECT C.cCtaCod Cuenta,Ltrim(Per.cPersNombre) Aval,C.nMontoApr Desembolso, " _
        '    & "         C.nSaldoCap Saldo,C.dFecVig Vigencia, " _
        '    & "         Isnull( Ti.SKTit,0) MontoEnMN, " _
        '    & "         (Case When SubString(C.cCtaCod,9,1)='1' Then 'MN' Else 'ME' End) Moneda, " _
        '    & "         (Case When SubString(C.cCtaCod,9,1)='1' Then Ti.SKTit Else 0 End) Cuenta841410, " _
        '    & "         (Case When SubString(C.cCtaCod,9,1)='2' Then Ti.SKTit * " & pnTipCambio & " Else 0 End) Cuenta842410, " _
        '    & "         Substring(C.cCtaCod,4,2) Agencia "
        'psSql = psSql & " " _
        '    & " FROM " & "DbConsolidada..CreditoConsol C " _
        '    & " INNER JOIN " & "DbConsolidada..ProductoPersonaConsol PP ON PP.cCtaCod = C.cCtaCod " _
        '    & " INNER JOIN " & "Persona Per ON Per.cPersCod = PP.cPersCod " _
        '    & " LEFT JOIN " & "DbConsolidada..RCDMaestroPersona RMP ON RMP.cPersCod = PP.cPersCod " _
        '    & " INNER JOIN Agencias Age ON Age.cAgeCod=SUBSTRING(C.cCtaCod,4,2) " _
        '    & " INNER JOIN (Select ca.cCtaCod CreditoTit, ca.cPersCod as Titular, ca.cCalGen as CalificaTit, ca.nMontoApr SKTit " _
        '    & "            From ColocCalifProv ca WHERE ca.nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) ) Ti " _
        '    & " ON Ti.CreditoTit = C.cCtaCod "
        
        '& " inner join " & psServConsol & "Rcavc2007110101 Ant On Ant.cCuentaTit=C.cCtaCod "
        'By Capi 1312007 solo por este mes, luego sacar
        'By Capi 13122007 se modifico para que jale el monto aprobado en vez del saldo de capital ca.nSaldoCap por ca.nMontoApr
        'ARCV 16-06-2007 : Comentar para el proximo mes
        'psSql = psSql & " LEFT JOIN RelClientes RCLI ON Ti.Titular =  RCLI.cPersCod " 'ARCV 16-06-2007
        '----------
        'psSql = psSql & " WHERE C.nPrdEstado in( " & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor _
        '    & "                       ," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor _
        '    & "                       ," & gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov _
        '    & "                        ) " _
        '    & "AND PP.nPrdPersRelac = 25 AND C.nSaldoCap > " & 0 _
        '    & " AND PP.cPersCod = (Select max(p1.cPersCod) From " & "DbConsolidada..ProductoPersonaConsol p1 " _
        '    & "                where p1.cCtacod = c.cCtaCod and p1.nPrdPersRelac = 25  )  Order By Aval,Vigencia"
        
        psSql = "exec stp_sel_ResumenGarantiaAvales " & pnTipCambio
        'END JUEZ ********************************************************************************************
   End If
   Set CargaEstadGarantias = dbConec.CargaRecordSet(psSql)
   dbConec.CierraConexion
   Set dbConec = Nothing
Exit Function
CargaEstadGarantiasErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:CargaEstadGarantias Method")
End Function
Public Function AjusteReclasificaGarantia(psFecha As String, pnMoneda As Integer, psClase As String, pnTipCambio As Currency) As Recordset
Dim sqlGar As String
Dim lsErrores As String
Dim lsNomSer As String
Dim oPrevio As PrevioFinan.clsPrevioFinan
Set oPrevio = New PrevioFinan.clsPrevioFinan
 
On Error GoTo AjusteReclasificaGarantiaErr
   Set dbConec = New DConecta
   dbConec.AbreConexion
   If pnBitCentral = True Then
'        lsNomSer = sServidorConsolidada
'        'sqlGar = FormaSelectGarantias(pnMoneda, "01", "01", "01", "", psClase, pnTipCambio, psFecha, lsNomSer)
'        sqlGar = ""
'        'Hipoteca 1 OK
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "", "02", "01", " 1,19 ", psClase, pnTipCambio, psFecha, lsNomSer, True, True) & " UNION ALL "
'        'Vehicular 2 OK
'        'Modificado por DAOR 20080711, De acuerdo a Acta Nº 149-2008/TI-D**********
'        'sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "", "09", "0902", "2", psClase, pnTipCambio, psFecha, lsNomSer, , True) & " UNION ALL "
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "", "09", "0902", "2,11,12,32,35,36", psClase, pnTipCambio, psFecha, lsNomSer, , True) & " UNION ALL " '*** PEAC 20110219 - SE AGREGO 32 (POCO A POCO AHORRO)
'        '***************************************************************************
'        'Industrial 3 OK
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "", "09", "0901", "3", psClase, pnTipCambio, psFecha, lsNomSer, , True) & " UNION ALL "
'        'Agricola 4 OK
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "", "09", "0202", "4", psClase, pnTipCambio, psFecha, lsNomSer, , True) & " UNION ALL "
'
'        '**Modificado por DAOR 20080410, según especificaciones de LEMA********************
'        ''Deposito 6 OK
'        'sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "", "04", "01", "6", psClase, pnTipCambio, psFecha, lsNomSer, , True) & " UNION ALL "
'        'Joyas 7 OK
'        'sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "", "02", "0601", "7", psClase, pnTipCambio, psFecha, lsNomSer, , True, True) & " UNION ALL "
'
'        'Deposito Plazo Fijo 6 y CTS 9
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "", "04", "0102", "6,9", psClase, pnTipCambio, psFecha, lsNomSer, , True) & " UNION ALL "
'        'Joyas 7 OK
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "", "02", "06", "7", psClase, pnTipCambio, psFecha, lsNomSer, , True, True) & " UNION ALL "
'        '**********************************************************************************
'
'        'Otras 10 autotizacion de desc  12 maq y equip 13 titulos 14 otros 15 otras garan ref 22 otras garan no ref 24 --FALTA AGREGAR DATOS
'        'ALPA 20081209******************************************************************************
'        'sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "", "09", "0902", " 10,12,13,14,15,22,23,24,25,26,27,28,29,30,18,34", psClase, pnTipCambio, psFecha, lsNomSer, True, True)
'        'sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "", "09", "0902", " 10,13,14,15,22,23,24,25,26,27,28,29,30,18,34", psClase, pnTipCambio, psFecha, lsNomSer, True, True)
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "", "09", "0902", " 10,13,14,15,22,23,24,25,26,27,28,29,30,18,34", psClase, pnTipCambio, psFecha, lsNomSer, True, True) & " UNION ALL "
'        '*******************************************************************************************
'        'ALPA20130722********************Para Garantias Leasing
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "", "02", "2901", " 38", psClase, pnTipCambio, psFecha, lsNomSer, True, True)
        '********************************
        psSql = "EXEC stp_sel_AjusteReclasificacionGarantias '" & psFecha & "'," & pnMoneda & "," & pnTipCambio 'EJVG20160505
   Else
'        'lsNomSer = dbConec.StringServidorRemoto("07", "03", True)
'        sqlGar = ""
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "04", "01", "01", "nDeposito", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "02", "01", "01", "nHipoteca", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "02", "06", "01", "nJoyas", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "02", "08", "01", "nIndustrial", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "02", "09", "01", "nVehicular", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "02", "11", "01", "nAgricola", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "02", "13", "01", "nGlobalFlo", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "05", "02", "01", "nCartaFianza", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "05", "02", "01", "nPersonal", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
'        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "05", "01", "01", " (nFianzaSolida+nAutorizacion+nOtrosG)", psClase, pnTipCambio, psFecha, lsNomSer)
        psSql = "EXEC stp_sel_AjusteReclasificacionGarantias '" & psFecha & "'," & pnMoneda & "," & pnTipCambio 'EJVG20160505
   End If
   'ALPA 20090227****************************************************
'   psSql = "SELECT cred.cCtaContCod cta1, SUM(ISNULL(cred.nSaldo,0)) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
'        & "FROM ( " & sqlGar & " ) Cred FULL OUTER JOIN " _
'        & "     ( SELECT cCtaContCod, nCtaSaldoImporte " _
'        & "       FROM CtaSaldo cs WHERE ISNULL(cs.nCtaSaldoImporte,0) <> 0 and cs.cCtaContCod like '" & psClase & pnMoneda & "4%'  and cs.dCtaSaldoFecha = " _
'        & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and substring(cs1.cCtaContCod,5,2) <>'10' and cs1.dCtaSaldoFecha <= '" & psFecha & "' ) " _
'        & "     ) cs on cs.cCtaContCod = cred.cCtaContCod " _
'        & "GROUP BY cred.cCtaContCod, cs.cCtaContCod, cs.nCtaSaldoImporte ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod) "
'   Set AjusteReclasificaGarantia = dbConec.CargaRecordSet(psSql)
'    psSql = "SELECT cred.cCtaContCod cta1, SUM(ISNULL(cred.nSaldo,0)) nSaldo, cs.cCtaContCod cta2, "
'    psSql = psSql & "ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte "
'    psSql = psSql & "FROM ( " & sqlGar & " ) Cred FULL OUTER JOIN "
'    psSql = psSql & "     ( SELECT cCtaContCod, "
'    If pnMoneda = 1 Then
'        psSql = psSql & " nCtaSaldoImporte "
'    Else
'        psSql = psSql & " nCtaSaldoImporte / " & IIf(pnTipCambio = 0, 1, pnTipCambio) & " nCtaSaldoImporte "
'    End If
'    psSql = psSql & "       FROM CtaSaldo cs WHERE ISNULL(cs.nCtaSaldoImporte,0) <> 0 and cs.cCtaContCod not in ('84" & pnMoneda & "4090903') and cs.cCtaContCod like '" & psClase & pnMoneda & "4%'  and cs.dCtaSaldoFecha = "
'    psSql = psSql & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and substring(cs1.cCtaContCod,5,2) <>'10' and cs1.dCtaSaldoFecha <= '" & psFecha & "' ) "
'    psSql = psSql & "     ) cs on cs.cCtaContCod = cred.cCtaContCod "
'    psSql = psSql & "GROUP BY cred.cCtaContCod, cs.cCtaContCod, cs.nCtaSaldoImporte ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod) "
   Set AjusteReclasificaGarantia = dbConec.CargaRecordSet(psSql)
   '******************************************************************
   lsErrores = ""
   Dim lrVerifica As New ADODB.Recordset
   If Not (AjusteReclasificaGarantia.BOF Or AjusteReclasificaGarantia.EOF) Then
        Do While Not AjusteReclasificaGarantia.EOF
            psSql = "Select count(*) as Existe FROM CTACONT Where cCtaContCod ='" & AjusteReclasificaGarantia!Cta1 & "' "
            Set lrVerifica = dbConec.CargaRecordSet(psSql)
            If lrVerifica!Existe = 0 Then
                If Not IsNull(AjusteReclasificaGarantia!Cta1) Then
                    'MsgBox "No existe Cuenta Contable " & AjusteReclasificaGarantia!cta1, vbCritical, "Aviso"
                    lsErrores = lsErrores & AjusteReclasificaGarantia!Cta1 & oImpresora.gPrnSaltoLinea
                End If
            End If
            AjusteReclasificaGarantia.MoveNext
            lrVerifica.Close
        Loop
        AjusteReclasificaGarantia.MoveFirst
        If lsErrores <> "" Then
            oPrevio.Show "Lista de Ctas No existentes : " & oImpresora.gPrnSaltoLinea & lsErrores, "", True
        End If
   End If
   
   Set dbConec = Nothing
Exit Function
AjusteReclasificaGarantiaErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:AjusteReclasificaGarantia Method")
End Function

Public Function AjusteReclasificaGarantiaAnt(psFecha As String, pnMoneda As Integer, psClase As String, pnTipCambio As Currency) As Recordset
Dim sqlGar As String
Dim lsNomSer As String
On Error GoTo AjusteReclasificaGarantiaErr
   Set dbConec = New DConecta
   dbConec.AbreConexion
   If pnBitCentral = True Then
        lsNomSer = sServidorConsolidada
        sqlGar = FormaSelectGarantias(pnMoneda, "01", "01", "01", "", psClase, pnTipCambio, psFecha, lsNomSer)
   Else
        'lsNomSer = dbConec.StringServidorRemoto("07", "03", True)
        sqlGar = ""
        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "01", "01", "01", "nDeposito", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "02", "01", "01", "nHipoteca", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "02", "06", "01", "nJoyas", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "02", "08", "01", "nIndustrial", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "02", "09", "01", "nVehicular", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "02", "11", "01", "nAgricola", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "02", "13", "01", "nGlobalFlo", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "02", "17", "01", "nCartaFianza", psClase, pnTipCambio, psFecha, lsNomSer) & " UNION ALL "
        sqlGar = sqlGar & FormaSelectGarantias(pnMoneda, "09", "01", "01", " (nFianzaSolida+nPersonal+nAutorizacion+nOtrosG)", psClase, pnTipCambio, psFecha, lsNomSer)
   End If

   psSql = "SELECT cred.cCtaContCod cta1, SUM(ISNULL(cred.nSaldo,0)) nSaldo, cs.cCtaContCod cta2, ISNULL(cs.nCtaSaldoImporte,0) nCtaSaldoImporte " _
        & "FROM ( " & sqlGar & " ) Cred FULL OUTER JOIN " _
        & "     ( SELECT cCtaContCod, nCtaSaldoImporte " _
        & "       FROM CtaSaldo cs WHERE ISNULL(cs.nCtaSaldoImporte,0) <> 0 and cs.cCtaContCod like '" & psClase & pnMoneda & "4%' and cs.dCtaSaldoFecha = " _
        & "           (SELECT max(dCtaSaldoFecha) FROM CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFecha & "' ) " _
        & "     ) cs on cs.cCtaContCod = cred.cCtaContCod " _
        & "GROUP BY cred.cCtaContCod, cs.cCtaContCod, cs.nCtaSaldoImporte ORDER BY ISNULL(cs.cCtaContCod,cred.cCtaContCod) "
   Set AjusteReclasificaGarantiaAnt = dbConec.CargaRecordSet(psSql)
   Set dbConec = Nothing
Exit Function
AjusteReclasificaGarantiaErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:AjusteReclasificaGarantia Method")
End Function

Private Function FormaSelectGarantias(pnMoneda As Integer, psPrefe As String, psTipo As String, psNaci As String, psConcepto As String, psClase As String, pnTipCambio As Currency, psFecha As String, psNomSer As String, Optional pbHiptecarios As Boolean = False, Optional pbFormaConsolidada As Boolean = False, Optional pbAgencia As Boolean = False) As String

If pnBitCentral = True Then
   If sCodEmp = "112" Or sCodEmp = "108" Or sCodEmp = "109" Then 'TRUJILLO
        If psNaci = "" Then
            If Not pbFormaConsolidada Then
                FormaSelectGarantias = "select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "' cCtaContCod, " _
                & "     ROUND(" & psConcepto & " * " & IIf(pnMoneda = 1, 1, pnTipCambio) & ", 2) nSaldo " _
                & "FROM " & psNomSer & "CredGarantias WHERE nMoneda = '" & pnMoneda & "' and " & psConcepto & " <> 0 and datediff(d,dFecha,'" & psFecha & "') = 0 "
            Else
                '& "     ROUND( (nMonto / (  (Case isnUll(CC.nSaldoCap,1) when 0 then 1 else isnUll(CC.nSaldoCap,1) end )  / (Case isNull(D.SaldoCap,1) when 0 then 1 else isNull(D.SaldoCap,1) end  ) )  )  * " & IIf(pnMoneda = 1, 1, pnTipCambio) & ", 2) nSaldo " _

                FormaSelectGarantias = " Select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'0'+SUBSTRING(CC.cCtaCod,6,1) + (SELECT LTRIM(RTRIM(cEquivAsiento)) From AsientoProdEquiv AE Where cEquivTpo ='SC' and AE.cProducto = SUBSTRING(CC.cCtaCod,6,3)) + (Select cCtaCont From dbconsolidada..coloclineacreditoequiv a where a.clineacred = LEFT(CC.cLineaCred,4)) + SUBSTRING(CC.cCtaCod,4,2)  cCtaContCod, " _
                & "        ROUND(nMonto*ROUND( CASE WHEN SUBSTRING(CC.CCTACOD,9,1)='1' THEN CC.nSaldoCap ELSE ROUND(CC.nSaldoCap*" & pnTipCambio & ",2) END / (CASE WHEN ISNULL(D.SaldoCap,0)=0 THEN 1 ELSE ISNULL(D.SaldoCap,0) END), 2)* " & IIf(pnMoneda = 1, 1, pnTipCambio) & ",2)  as nSaldo " _
                & " FROM " & psNomSer & "CredGarantiasConsol CG " _
                & " Inner Join " & psNomSer & "ProductoPersonaConsol PP On CG.cCtaCod = PP.cCtaCod" _
                & " Inner Join " & psNomSer & "CreditoConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 20" _
                & " Inner Join (" _
                & "       select  GC.cPersCod, Sum(  Case Substring(B.cctacod,9,1) When '1' then B.nSaldoCap else B.nSaldoCap * " & pnTipCambio & " end) SaldoCap, Count(*) Num" _
                & "       from    " & psNomSer & "ProductoPersonaConsol GC" _
                & "       INNER jOIN " & psNomSer & "creditoconsol B On GC.cCtaCod = B.cCtaCod and GC.nPrdPersRelac = 20" _
                & "       group by GC.cPersCod) As D On PP.cPersCod = D.cPersCod " _
                & " WHERE  nMoneda = '" & pnMoneda & "' And cTipoGarant In (" & psConcepto & ") And datediff(d,CG.dFecha,'" & psFecha & "') = 0 "
                
                
            End If
        Else
            If pbHiptecarios Then
                If Not pbFormaConsolidada Then
                    FormaSelectGarantias = "select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "'+'0'+SUBSTRING(cCtaCod,6,1) + " _
                    & "CASE WHEN SUBSTRING(cCtaCod,6,3) = '" & Producto.gColConsuPrendario & "' THEN '13' " _
                    & "     WHEN SUBSTRING(cCtaCod,6,3) = '" & Producto.gColConsuPrestAdm & "' THEN '20' " _
                    & "     ELSE '01' END + " _
                    & "     ISNULL(CASE WHEN cFF = 'A' THEN '10' WHEN cFF = 'B' THEN '11' WHEN cFF = 'C' THEN '12' ELSE cFF END,'01') + SUBSTRING(cCtaCod,4,2) cCtaContCod, " _
                    & "     ROUND(" & psConcepto & " * " & IIf(pnMoneda = 1, 1, pnTipCambio) & ", 2) nSaldo " _
                    & "FROM " & psNomSer & "CredGarantias WHERE nMoneda = '" & pnMoneda & "' and " & psConcepto & " <> 0 and datediff(d,dFecha,'" & psFecha & "') = 0 "
                Else
                    '**Modificado por DAOR 20080410 ********************************************
                    'FormaSelectGarantias = "select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "' " _
                    '& "     cCtaContCod, " _
                    '& "     ROUND(nMonto*ROUND( CASE WHEN SUBSTRING(CC.CCTACOD,9,1)='1' THEN CC.nSaldoCap ELSE ROUND(CC.nSaldoCap*" & pnTipCambio & ",2) END / (CASE WHEN ISNULL(D.SaldoCap,0)=0 THEN 1 ELSE ISNULL(D.SaldoCap,0) END), 2),2) * " & IIf(pnMoneda = 1, 1, pnTipCambio) & " as nSaldo " _
                    '& " FROM " & psNomSer & "CredGarantiasConsol CG " _
                    '& "     Inner Join " & psNomSer & "ProductoPersonaConsol PP On CG.cCtaCod = PP.cCtaCod" _
                    '& "     Inner Join " & psNomSer & "CreditoConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 20" _
                    '& "     Inner Join (" _
                    '& "           select  GC.cPersCod, Sum(  Case Substring(B.cctacod,9,1) When '1' then B.nSaldoCap else B.nSaldoCap * " & pnTipCambio & " end) SaldoCap, Count(*) Num" _
                    '& "         from    " & psNomSer & "ProductoPersonaConsol GC" _
                    '& "         INNER jOIN " & psNomSer & "creditoconsol B On GC.cCtaCod = B.cCtaCod and GC.nPrdPersRelac = 20" _
                    '& "         group by GC.cPersCod) As D On PP.cPersCod = D.cPersCod " _
                    '& " WHERE nMoneda = '" & pnMoneda & "' And cTipoGarant In (" & psConcepto & ") and datediff(d,CG.dFecha,'" & psFecha & "') = 0 "
                    
                    'ALPA 20151219**************************************************************************************************
'                    FormaSelectGarantias = "select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "' " _
'                    & "     cCtaContCod, round(CG.nMonto,2) as nSaldo " _
'                    & " from " & psNomSer & "CredGarantiasConsol CG " _
'                    & "     inner join " & psNomSer & "ProductoPersonaConsol PP On CG.cCtaCod = PP.cCtaCod" _
'                    & "     inner join " & psNomSer & "CreditoConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 20" _
'                    & " where  CG.nMoneda = '" & pnMoneda & "' and CG.cTipoGarant In (" & psConcepto & ") and datediff(d,CG.dFecha,'" & psFecha & "') = 0 "
                    FormaSelectGarantias = "select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "' " _
                    & "     cCtaContCod, round(CG.nMonto,2) as nSaldo " _
                    & " from " & psNomSer & "CredGarantiasConsol CG " _
                    & "     inner join " & psNomSer & "ProductoPersonaConsol PP On CG.cCtaCod = PP.cCtaCod" _
                    & "     inner join " & psNomSer & "CreditoConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 20" _
                    & " where  CG.nMoneda = '" & pnMoneda & "' and (CG.cCtaContCod like '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "%') and datediff(d,CG.dFecha,'" & psFecha & "') = 0 "
                    '****************************************************************************
                    
                End If
            Else
                If Not pbFormaConsolidada Then
                    FormaSelectGarantias = "select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "'+'0'+SUBSTRING(cCtaCod,6,1)+" _
                    & "CASE WHEN SUBSTRING(cCtaCod,6,3) = '" & Producto.gColConsuPrendario & "' THEN '13' " _
                    & "     WHEN SUBSTRING(cCtaCod,6,3) = '" & Producto.gColConsuPrestAdm & "' THEN '20' " _
                    & "     ELSE SUBSTRING(cCtaCod,7,2) END + " _
                    & "     ISNULL(CASE WHEN cFF = 'A' THEN '10' WHEN cFF = 'B' THEN '11' WHEN cFF = 'C' THEN '12' ELSE cFF END,'01') + SUBSTRING(cCtaCod,4,2) cCtaContCod, " _
                    & "     ROUND(" & psConcepto & " * " & IIf(pnMoneda = 1, 1, pnTipCambio) & ", 2) nSaldo " _
                    & "FROM " & psNomSer & "CredGarantias WHERE nMoneda = '" & pnMoneda & "' and " & psConcepto & " <> 0 and datediff(d,dFecha,'" & psFecha & "') = 0 "
                Else
                    '**Modificado por DAOR 20080410************************************************
                    'FormaSelectGarantias = "select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "' " _
                    '& "  " & IIf(pbAgencia, "  + SUBSTRING(cc.cCtaCod,4,2) ", "") & " cCtaContCod, " _
                    '& "         ROUND(nMonto*ROUND( CASE WHEN SUBSTRING(CC.CCTACOD,9,1)='1' THEN CC.nSaldoCap ELSE ROUND(CC.nSaldoCap*" & pnTipCambio & ",2) END / (CASE WHEN ISNULL(D.SaldoCap,0)=0 THEN 1 ELSE ISNULL(D.SaldoCap,0) END), 2)* " & IIf(pnMoneda = 1, 1, pnTipCambio) & ",2) as nSaldo " _
                    '& "FROM " & psNomSer & "CredGarantiasConsol CG " _
                    '& " Inner Join " & psNomSer & "ProductoPersonaConsol PP On CG.cCtaCod = PP.cCtaCod" _
                    '& " Inner Join " & psNomSer & "CreditoConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 20" _
                    '& " Inner Join (" _
                    '& "       select  GC.cPersCod, Sum(  Case Substring(B.cctacod,9,1) When '1' then B.nSaldoCap else B.nSaldoCap * " & pnTipCambio & " end) SaldoCap, Count(*) Num" _
                    '& "       from    " & psNomSer & "ProductoPersonaConsol GC" _
                    '& "       INNER jOIN " & psNomSer & "creditoconsol B On GC.cCtaCod = B.cCtaCod and GC.nPrdPersRelac = 20" _
                    '& "       group by GC.cPersCod) As D On PP.cPersCod = D.cPersCod " _
                    '& " WHERE nMoneda = '" & pnMoneda & "' And cTipoGarant In (" & psConcepto & ") and datediff(d,CG.dFecha,'" & psFecha & "') = 0 "
                    
                    If psConcepto = "7" Then 'Pignoraticio
'                        FormaSelectGarantias = "select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "' + case when CC.nDiasAtraso <= 30 then '01' else '02' end " _
'                        & "  " & IIf(pbAgencia, "  + SUBSTRING(CC.cCtaCod,4,2) ", "") & " cCtaContCod, " _
'                        & "         round(CG.nMonto,2) as nSaldo " _
'                        & "from " & psNomSer & "CredGarantiasConsol CG " _
'                        & " inner join " & psNomSer & "ProductoPersonaConsol PP On CG.cCtaCod = PP.cCtaCod" _
'                        & " inner join " & psNomSer & "CreditoConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 20" _
'                        & " where CG.nMoneda = '" & pnMoneda & "' And CG.cTipoGarant In (" & psConcepto & ") and datediff(d,CG.dFecha,'" & psFecha & "') = 0 "
                        'ALPA 20151219***********************************************
                        FormaSelectGarantias = "select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "' + case when CC.nDiasAtraso <= 30 then '01' else '02' end " _
                        & "  " & IIf(pbAgencia, "  + SUBSTRING(CC.cCtaCod,4,2) ", "") & " cCtaContCod, " _
                        & "         round(CG.nMonto,2) as nSaldo " _
                        & "from " & psNomSer & "CredGarantiasConsol CG " _
                        & " inner join " & psNomSer & "ProductoPersonaConsol PP On CG.cCtaCod = PP.cCtaCod" _
                        & " inner join " & psNomSer & "CreditoConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 20" _
                        & " where CG.nMoneda = '" & pnMoneda & "' And CG.cTipoGarant In (" & psConcepto & ") " _
                        & " and CG.cCtaContCod like '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "%') and datediff(d,CG.dFecha,'" & psFecha & "') = 0 "
                    Else
'                        FormaSelectGarantias = "select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "' " _
'                        & "  " & IIf(pbAgencia, "  + SUBSTRING(CC.cCtaCod,4,2) ", "") & " cCtaContCod, " _
'                        & "         round(CG.nMonto,2) as nSaldo " _
'                        & "from " & psNomSer & "CredGarantiasConsol CG " _
'                        & " inner join " & psNomSer & "ProductoPersonaConsol PP On CG.cCtaCod = PP.cCtaCod" _
'                        & " inner join " & psNomSer & "CreditoConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 20" _
'                        & " where CG.nMoneda = '" & pnMoneda & "' And CG.cTipoGarant In (" & psConcepto & ") and datediff(d,CG.dFecha,'" & psFecha & "') = 0 "
                        FormaSelectGarantias = "select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "' " _
                        & "  " & IIf(pbAgencia, "  + SUBSTRING(CC.cCtaCod,4,2) ", "") & " cCtaContCod, " _
                        & "         round(CG.nMonto,2) as nSaldo " _
                        & "from " & psNomSer & "CredGarantiasConsol CG " _
                        & " inner join " & psNomSer & "ProductoPersonaConsol PP On CG.cCtaCod = PP.cCtaCod" _
                        & " inner join " & psNomSer & "CreditoConsol CC On CC.cCtaCod = PP.cCtaCod And PP.nPrdPersRelac = 20" _
                        & " where CG.nMoneda = '" & pnMoneda & "' And CG.cTipoGarant In (" & psConcepto & ") " _
                        & " and (CG.cCtaContCod like '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "%') and datediff(d,CG.dFecha,'" & psFecha & "') = 0 "
                    End If
                    '*********************************************************************************

                End If
            End If
        End If
   End If
   If sCodEmp = "102" Then    'LIMA
        FormaSelectGarantias = "SELECT '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+ right('0'+rtrim(cg.cTipoGarant),2) +'" & psNaci & "'+'0'+SUBSTRING(cCtaCod,3,1)+" _
        & "CASE WHEN SUBSTRING(cCtaCod,6,3) = '" & Producto.gColConsuPrendario & "' THEN '13' " _
        & "     WHEN SUBSTRING(cCtaCod,6,3) = '" & Producto.gColConsuPrestAdm & "' THEN '20' " _
        & "     ELSE SUBSTRING(cCtaCod,7,2) END + " _
        & "     ISNULL(cFF,'01') + SUBSTRING(cCtaCod,4,2) cCtaContCod, " _
        & "     ROUND( nMonto * " & IIf(pnMoneda = 1, 1, pnTipCambio) & ", 2) nSaldo " _
        & "FROM " & psNomSer & "CredGarantiasConsol cg JOIN Constante c ON cg.cTipoGarant = nConsValor and c.nConsCod = " & gPersGarantia & " " _
        & "WHERE nMoneda = '" & pnMoneda & "' and datediff(d,dFecha,'" & psFecha & "') = 0 "
   End If
Else
    FormaSelectGarantias = "select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "'+'0'+SUBSTRING(cCodCta,3,1)+" _
        & "CASE WHEN SUBSTRING(cCodCta,3,3) = '305' THEN '13' " _
        & "     WHEN SUBSTRING(cCodCta,3,3) = '320' THEN '20' " _
        & "     ELSE SUBSTRING(cCodCta,4,2) END + " _
        & "     ISNULL(CASE WHEN cFF = 'A' THEN '10' WHEN cFF = 'B' THEN '11' WHEN cFF = 'C' THEN '12' ELSE '0'+cFF END,'01')+ LEFT(cCodCta,2) cCtaContCod, " _
        & "     ROUND(" & psConcepto & " * " & IIf(pnMoneda = 1, 1, pnTipCambio) & ", 2) nSaldo " _
        & "FROM " & psNomSer & "CredGarantias WHERE cMoneda = '" & pnMoneda & "' and " & psConcepto & " <> 0 and datediff(d,dFecha,'" & psFecha & "') = 0 "
End If
End Function

Private Function FormaSelectGarantiasAnt(pnMoneda As Integer, psPrefe As String, psTipo As String, psNaci As String, psConcepto As String, psClase As String, pnTipCambio As Currency, psFecha As String, psNomSer As String) As String

If pnBitCentral = True Then
   If sCodEmp = "112" Or sCodEmp = "108" Then 'TRUJILLO
        FormaSelectGarantiasAnt = "select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "'+'0'+SUBSTRING(cCtaCod,3,1)+" _
        & "CASE WHEN SUBSTRING(cCtaCod,6,3) = '" & Producto.gColConsuPrendario & "' THEN '13' " _
        & "     WHEN SUBSTRING(cCtaCod,6,3) = '" & Producto.gColConsuPrestAdm & "' THEN '20' " _
        & "     ELSE SUBSTRING(cCtaCod,7,2) END + " _
        & "     ISNULL(CASE WHEN cFF = 'A' THEN '10' WHEN cFF = 'B' THEN '11' WHEN cFF = 'C' THEN '12' ELSE '0'+cFF END,'01') + SUBSTRING(cCtaCod,4,2) cCtaContCod, " _
        & "     ROUND(" & psConcepto & " * " & IIf(pnMoneda = 1, 1, pnTipCambio) & ", 2) nSaldo " _
        & "FROM " & psNomSer & "CredGarantias WHERE nMoneda = '" & pnMoneda & "' and " & psConcepto & " <> 0 and datediff(d,dFecha,'" & psFecha & "') = 0 "
   End If
   If sCodEmp = "102" Or sCodEmp = "108" Then    'LIMA
        FormaSelectGarantiasAnt = "SELECT '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+ right('0'+rtrim(cg.cTipoGarant),2) +'" & psNaci & "'+'0'+SUBSTRING(cCtaCod,3,1)+" _
        & "CASE WHEN SUBSTRING(cCtaCod,6,3) = '" & Producto.gColConsuPrendario & "' THEN '13' " _
        & "     WHEN SUBSTRING(cCtaCod,6,3) = '" & Producto.gColConsuPrestAdm & "' THEN '20' " _
        & "     ELSE SUBSTRING(cCtaCod,7,2) END + " _
        & "     ISNULL(cFF,'01') + SUBSTRING(cCtaCod,4,2) cCtaContCod, " _
        & "     ROUND( nMonto * " & IIf(pnMoneda = 1, 1, pnTipCambio) & ", 2) nSaldo " _
        & "FROM " & psNomSer & "CredGarantiasConsol cg JOIN Constante c ON cg.cTipoGarant = nConsValor and c.nConsCod = " & gPersGarantia & " " _
        & "WHERE nMoneda = '" & pnMoneda & "' and datediff(d,dFecha,'" & psFecha & "') = 0 "
   End If
Else
    FormaSelectGarantiasAnt = "select '" & psClase & pnMoneda & "4'+'" & psPrefe & "'+'" & psTipo & "'+'" & psNaci & "'+'0'+SUBSTRING(cCodCta,3,1)+" _
        & "CASE WHEN SUBSTRING(cCodCta,3,3) = '305' THEN '13' " _
        & "     WHEN SUBSTRING(cCodCta,3,3) = '320' THEN '20' " _
        & "     ELSE SUBSTRING(cCodCta,4,2) END + " _
        & "     ISNULL(CASE WHEN cFF = 'A' THEN '10' WHEN cFF = 'B' THEN '11' WHEN cFF = 'C' THEN '12' ELSE '0'+cFF END,'01')+ LEFT(cCodCta,2) cCtaContCod, " _
        & "     ROUND(" & psConcepto & " * " & IIf(pnMoneda = 1, 1, pnTipCambio) & ", 2) nSaldo " _
        & "FROM " & psNomSer & "CredGarantias WHERE cMoneda = '" & pnMoneda & "' and " & psConcepto & " <> 0 and datediff(d,dFecha,'" & psFecha & "') = 0 "
End If
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    pnBitCentral = IIf(oConst.LeeConstSistema(gConstSistBitCentral) = "1", True, False)
    
    Set oImp = New DImpresoras
        oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    Set oImp = Nothing

    Dim oIni As New ClasIni
        sCentralCom = oIni.BaseComunes
    Set oIni = Nothing
    
    sCodEmp = oConst.LeeConstSistema(gConstSistCodCMAC)
    sServidorConsolidada = oConst.LeeConstSistema(gConstSistServCentralRiesgos)
 
    Set oConst = Nothing
      
End Sub
'ALPA 20090525**************************
Function DevuelveMesAnio(ByVal pdFecha As Date, Optional ByVal nTipo As Integer = 1) As String
    'nTipo 1 Mes
    'nTipo 2 Año
    
    Dim sMes As String
    Dim sAnio As String
    sMes = CStr(IIf(CInt(Format(pdFecha, "MM")) = 1, 12, CInt(Format(pdFecha, "MM")) - 1))
    If nTipo = 1 Then
        
        DevuelveMesAnio = IIf(Len(sMes) = 1, "0" & sMes, sMes)
        Exit Function
    Else
        If sMes = "12" Then
            sAnio = CStr(CInt(Format(pdFecha, "YYYY")) - 1)
        Else
            sAnio = Format(pdFecha, "YYYY")
        End If
        DevuelveMesAnio = sAnio
        Exit Function
    End If
    DevuelveMesAnio = ""
End Function
'***********************************
Public Function VerificarSumaProvisionProciclica() As Currency
Dim sCond As String
   On Error GoTo VerificarSumaProvisionProciclicaErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "select sum(isnull(nProvisionProciclica,0))nSuma from ColocCalifProv  "
      Set rs = dbConec.CargaRecordSet(psSql)
      If Not (rs.BOF Or rs.EOF) Then
        VerificarSumaProvisionProciclica = rs!nSuma
      Else
        VerificarSumaProvisionProciclica = 0
      End If
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
VerificarSumaProvisionProciclicaErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:CargaAjusteHistorico Method")
End Function
'ALPA 20120403****************************************************
Public Function GenerarAsientoFractal(pdFecha As Date, psOpecod As String, psCodUser As String, ByRef psMovNro As String, Optional ByVal pnTipoCambio As Currency = 0) As String
Dim sCond As String
   On Error GoTo GenerarAsientoFractalErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Dim rsA As ADODB.Recordset
   Set rsA = New ADODB.Recordset
      
   Set dbConec = New DConecta
   psMovNro = ""
   If dbConec.AbreConexion() Then
       
      psSql = "exec stp_sel_ValidaGeneraAsiento '" & Format(pdFecha, "YYYY/MM/DD") & "','" & psOpecod & "'"
      Set rs = dbConec.CargaRecordSet(psSql)
      If Not (rs.BOF Or rs.EOF) Then
          'ALPA 20140130
          'Comentado porque no paso el nuevo asiento de fractal
          psSql = "exec stp_ins_GeneraAsiento '" & Format(pdFecha, "YYYY/MM/DD") & "','" & psOpecod & "','" & psCodUser & "'"
          'psSql = "exec stp_ins_GeneraAsiento '" & Format(pdFecha, "YYYY/MM/DD") & "','" & psOpeCod & "','" & psCodUser & "'," & pnTipoCambio
          dbConec.CargaRecordSet (psSql)
          psSql = "exec stp_sel_GeneraAsiento '" & Format(pdFecha, "YYYY/MM/DD") & "','" & psOpecod & "','" & psCodUser & "'"
          Set rsA = dbConec.CargaRecordSet(psSql)
          If Not (rsA.BOF Or rsA.EOF) Then
            psMovNro = IIf(IsNull(rsA!cMovNro), "", rsA!cMovNro)
          End If
        GenerarAsientoFractal = "Asiento Generado correctamente"
      Else
        GenerarAsientoFractal = "No existen datos"
      End If
   
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
GenerarAsientoFractalErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:CargaAjusteHistorico Method")
End Function

Public Function ObtieneAsientoFractal(pdFecha As Date, psOpecod As String, ByRef psMovNro As String) As String
Dim sCond As String
   On Error GoTo ObtieneAsientoFractalErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Dim rsA As ADODB.Recordset
   Set rsA = New ADODB.Recordset
      
   Set dbConec = New DConecta
   psMovNro = ""
   If dbConec.AbreConexion() Then
       
      psSql = "exec stp_sel_ObtieneMovFractal '" & Format(pdFecha, "YYYY/MM/DD") & "','" & psOpecod & "'"
      Set rs = dbConec.CargaRecordSet(psSql)
      If Not (rs.BOF Or rs.EOF) Then
        psMovNro = IIf(IsNull(rs!cMovNro), "", rs!cMovNro)
        ObtieneAsientoFractal = "Asiento Generado correctamente"
      Else
        ObtieneAsientoFractal = "No existen datos"
      End If
   
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
ObtieneAsientoFractalErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:CargaAjusteHistorico Method")
End Function

'*******************************************************************
'ALPA 20121123******************************************************
Public Function PlanillaElectronicaVenta(pdFecha As Date, pnMoneda As String) As String
    Dim lsCadena, lsSql As String
    Dim oConecta As DConecta
    Dim lnNumeroLineas As Integer
    Dim lnSumaLineas As Integer
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Set oConecta = New DConecta
    lnNumeroLineas = 0
    lnSumaLineas = 0
    oConecta.AbreConexion
    lsSql = "stp_sel_PlaElecRegistroVenta '" & Format(pdFecha, "YYYY/MM/DD") & "','" & pnMoneda & "'"
    Set oRs = oConecta.CargaRecordSet(lsSql)
    If Not (oRs.BOF Or oRs.EOF) Then
        lsCadena = ""
        lnNumeroLineas = oRs.RecordCount
        Do While Not oRs.EOF
        lnSumaLineas = lnSumaLineas + 1
        lsCadena = lsCadena & oRs!Periodo_1
        lsCadena = lsCadena & "|" & oRs!Correlativo_2
        lsCadena = lsCadena & "|" & Format(oRs!FechaEmision_3, "DD/MM/YYYY")
        lsCadena = lsCadena & "|" & "" 'oRs!FechaVencimi_4
        lsCadena = lsCadena & "|" & oRs!TipoCdP_5
        lsCadena = lsCadena & "|" & oRs!Serie_6
        lsCadena = lsCadena & "|" & oRs!Dua_7
        lsCadena = lsCadena & "|" & oRs!Numero_8
        'lsCadena = lsCadena & "|" & oRs!dato_9
        lsCadena = lsCadena & "|" & oRs!TDI_10
        lsCadena = lsCadena & "|" & oRs!DI_11
        lsCadena = lsCadena & "|" & oRs!Nom_12
        lsCadena = lsCadena & "|" & oRs!BI_13
        lsCadena = lsCadena & "|" & oRs!IM_14
        lsCadena = lsCadena & "|" & oRs!BI_15
        lsCadena = lsCadena & "|" & oRs!IM_16
        lsCadena = lsCadena & "|" & oRs!BI_17
        lsCadena = lsCadena & "|" & oRs!IM_18
        lsCadena = lsCadena & "|" & oRs!VANG_19
        lsCadena = lsCadena & "|" & oRs!ICS_20
        lsCadena = lsCadena & "|" & oRs!OT_21
        lsCadena = lsCadena & "|" & oRs!ITASCP_22
        lsCadena = lsCadena & "|" & oRs!TC_23
        lsCadena = lsCadena & "|" & IIf(CStr(Format(IIf(Trim(oRs!FCPM_24) = "1901-01-01", "01/01/0001", oRs!FCPM_24), "DD/MM/YYYY")) = "01/01/2001", "01/01/0001", Format(oRs!FCPM_24, "DD/MM/YYYY"))
        lsCadena = lsCadena & "|" & oRs!TCM_25
        lsCadena = lsCadena & "|" & oRs!NSCM_26
        lsCadena = lsCadena & "|" & oRs!NCM_27
        
'        lsCadena = lsCadena & "|" & oRs!NCMSND_28
'        lsCadena = lsCadena & "|" & oRs!FED_29
'        lsCadena = lsCadena & "|" & oRs!NCOD_30
'        lsCadena = lsCadena & "|" & oRs!MCPSR_31
        lsCadena = lsCadena & "|" & oRs!ESTADO_32 & "|"
        If lnNumeroLineas <> lnSumaLineas Then
            lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        End If

        oRs.MoveNext
        Loop
    Else
        MsgBox "No existe datos para la PLE de Registro de Venta"
    End If
    oConecta.CierraConexion
    PlanillaElectronicaVenta = lsCadena
End Function
Public Function PlanillaElectronicaCompra(pdFecha As Date, pnMoneda As String) As String
    Dim lsCadena, lsSql As String
    Dim lnNumeroLineas As Integer
    Dim lnSumaLineas As Integer
    Dim oConecta As DConecta
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    Set oConecta = New DConecta
    lnNumeroLineas = 0
    lnSumaLineas = 0
    oConecta.AbreConexion
    lsSql = "stp_sel_PlaElecRegistroCompra '" & Format(pdFecha, "YYYY/MM/DD") & "'," & pnMoneda & ""
    Set oRs = oConecta.CargaRecordSet(lsSql)
    If Not (oRs.BOF Or oRs.EOF) Then
        lsCadena = ""
        lnNumeroLineas = oRs.RecordCount
        Do While Not oRs.EOF
        lnSumaLineas = lnSumaLineas + 1
        lsCadena = lsCadena & oRs!Periodo_1
        lsCadena = lsCadena & "|" & oRs!Correlativo2_
        lsCadena = lsCadena & "|" & Format(oRs!FechaEmision3_, "DD/MM/YYYY")
        lsCadena = lsCadena & "|" & IIf(oRs!TipoCdP5_ = "14", oRs!FechaCancela4_, "") '& Format(oRs!FechaCancela4_, "DD/MM/YYYY")
        lsCadena = lsCadena & "|" & oRs!TipoCdP5_
        lsCadena = lsCadena & "|" & oRs!Serie6_
        lsCadena = lsCadena & "|" & oRs!Dua7_
        lsCadena = lsCadena & "|" & oRs!Numero8_
        lsCadena = lsCadena & "|" & oRs!NumeroTotal9_
        lsCadena = lsCadena & "|" & oRs!TDP10_
        lsCadena = lsCadena & "|" & oRs!RUC11_
        lsCadena = lsCadena & "|" & oRs!PersNombre12_
        lsCadena = lsCadena & "|" & oRs!BaseImponible13_
        lsCadena = lsCadena & "|" & oRs!IGV14_
        lsCadena = lsCadena & "|" & oRs!PV215_
        lsCadena = lsCadena & "|" & oRs!IGV16_
        lsCadena = lsCadena & "|" & oRs!PV317_
        lsCadena = lsCadena & "|" & oRs!IGV318_
        lsCadena = lsCadena & "|" & oRs!VANG19_
        lsCadena = lsCadena & "|" & oRs!MISC20_
        lsCadena = lsCadena & "|" & oRs!Otros21_
        lsCadena = lsCadena & "|" & oRs!IT22_
        lsCadena = lsCadena & "|" & Format(oRs!TC23_, "0.##0")
        lsCadena = lsCadena & "|" & IIf(CStr(Format(IIf(Trim(oRs!FM24) = "1901-01-01", "01/01/0001", oRs!FM24), "DD/MM/YYYY")) = "01/01/2001", "01/01/0001", Format(oRs!FM24, "DD/MM/YYYY")) 'oRs!FM24
        lsCadena = lsCadena & "|" & oRs!TCM25_
        lsCadena = lsCadena & "|" & oRs!NS26_
        lsCadena = lsCadena & "|" & oRs!NDM27_
        lsCadena = lsCadena & "|" & oRs!NDND28_
        lsCadena = lsCadena & "|" & oRs!FECD29_
        lsCadena = lsCadena & "|" & oRs!NCD30_
        lsCadena = lsCadena & "|" & oRs!MCPSR31_
        lsCadena = lsCadena & "|" & oRs!Estado & "|"
        If lnNumeroLineas <> lnSumaLineas Then
            lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        End If
        oRs.MoveNext
        Loop
    End If
    oConecta.CierraConexion
    PlanillaElectronicaCompra = Trim(lsCadena)
End Function

'*** PEAC 20130111
Public Function AjusteIntDevCaptaciones(pnTipoAsiento As Integer, psFecha As String, pnMoneda As Integer, psClase As String, pnTipCambio As Currency, psCtaDebe As String, Optional pbAsiRevDev As Boolean = False, Optional pnTipoCambioVenta As Currency = 0, Optional pnTipoCambioCompra As Currency = 0) As Recordset
On Error GoTo AjusteIntDevCaptacionesErr
   
   Set dbConec = New DConecta
   dbConec.AbreConexion
 
    If pnBitCentral = True Then
                  
        psSql = "exec stp_sel_AjusteIntDevengadoSaldosAhorro '" & Trim(Str(pnMoneda)) & "','" & psFecha & "'," & pnTipCambio
                                        
    End If
      
   Set AjusteIntDevCaptaciones = dbConec.CargaRecordSet(psSql)
   Set dbConec = Nothing
Exit Function
AjusteIntDevCaptacionesErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:AjusteIntDevCaptaciones Method")
End Function
'PASIERS0142015***********************************************************************************
Public Function AjusteDepreciacionAdjudicados(psFecha As String) As ADODB.Recordset
    On Error GoTo AjusteDeprecAdjudicErr
    
    Set dbConec = New DConecta
    dbConec.AbreConexion
    
    If pnBitCentral Then
        psSql = "stp_sel_ObtieneDepreAdjudicxAjuste '" & psFecha & "'"
    End If
    
    Set AjusteDepreciacionAdjudicados = dbConec.CargaRecordSet(psSql)
    Set dbConec = Nothing
Exit Function
AjusteDeprecAdjudicErr:
    Call RaiseError(MyUnhandledError, "DAjusteCont:AjusteDepreciacionAdjudicados Method")
End Function
'END PASI
'PASI20160216 ERS0672015
Public Function AjusteParaGramosDeOroEnCustodia(psFecha As String) As ADODB.Recordset
    On Error GoTo AjusteParaGramosDeOroEnCustodiaERR
    
    Set dbConec = New DConecta
    dbConec.AbreConexion
    
    If pnBitCentral Then
        psSql = "stp_sel_ObtieneAjusteParaGramosDeOroEnCustodia '" & psFecha & "'"
    End If
    
    Set AjusteParaGramosDeOroEnCustodia = dbConec.CargaRecordSet(psSql)
    Set dbConec = Nothing
Exit Function
AjusteParaGramosDeOroEnCustodiaERR:
    Call RaiseError(MyUnhandledError, "DAjusteCont:AjusteParaGramosDeOroEnCustodia Method")
End Function
'END PASI
'PASI20160223 ERS0072016
Public Function AjusteIntDevengadoDet(ByVal psClaseCta As String, ByVal psMoneda As String, ByVal pdFechaFin As Date) As ADODB.Recordset
On Error GoTo AjusteIDError
Set dbConec = New DConecta
dbConec.AbreConexion
    psSql = "stp_sel_AjusteInteresesDevengadosDet '" & psClaseCta & "','" & psMoneda & "','" & Format(pdFechaFin, "yyyymmdd") & "'"
    Set AjusteIntDevengadoDet = dbConec.CargaRecordSet(psSql)
    Set dbConec = Nothing
Exit Function
AjusteIDError:
    Call RaiseError(MyUnhandledError, "DAjusteCont:AjusteIntDevengadoDet Method")
End Function
Public Function AjusteIntSuspensoDet(ByVal psClaseCta As String, ByVal psMoneda As String, ByVal pdFechaFin As Date) As ADODB.Recordset
On Error GoTo AjusteIDError
Set dbConec = New DConecta
dbConec.AbreConexion
    psSql = "stp_sel_AjusteInteresesSuspensoDet '" & psClaseCta & "','" & psMoneda & "','" & Format(pdFechaFin, "yyyymmdd") & "'"
    Set AjusteIntSuspensoDet = dbConec.CargaRecordSet(psSql)
    Set dbConec = Nothing
Exit Function
AjusteIDError:
    Call RaiseError(MyUnhandledError, "DAjusteCont:AjusteIntSuspensoDet Method")
End Function
Public Function AjusteRevIntDevengadoDet(ByVal psClaseCta As String, ByVal psMoneda As String, ByVal pdFechaFin As Date) As ADODB.Recordset
On Error GoTo AjusteIDError
Set dbConec = New DConecta
dbConec.AbreConexion
    psSql = "stp_sel_AjusteReversionInteresesDevengadosDet '" & psClaseCta & "','" & psMoneda & "','" & Format(pdFechaFin, "yyyymmdd") & "'"
    Set AjusteRevIntDevengadoDet = dbConec.CargaRecordSet(psSql)
    Set dbConec = Nothing
Exit Function
AjusteIDError:
    Call RaiseError(MyUnhandledError, "DAjusteCont:AjusteRevIntDevengadoDet Method")
End Function
Public Function AjusteProvisionCarteraDet(ByVal psClaseCta As String, ByVal psMoneda As String, ByVal pdFechaFin As Date) As ADODB.Recordset
On Error GoTo AjusteIDError
Set dbConec = New DConecta
dbConec.AbreConexion
    psSql = "stp_sel_AjusteProvisionCarteraDet '" & psClaseCta & "','" & psMoneda & "','" & Format(pdFechaFin, "yyyymmdd") & "'"
    Set AjusteProvisionCarteraDet = dbConec.CargaRecordSet(psSql)
    Set dbConec = Nothing
Exit Function
AjusteIDError:
    Call RaiseError(MyUnhandledError, "DAjusteCont:AjusteProvisionCarteraDet Method")
End Function
Public Function InteresDevengadoDetXCta(ByVal psCtaCod As String, ByVal pdFechaIni As Date, ByVal pdFechaFin As Date) As ADODB.Recordset
On Error GoTo RepError
    Set dbConec = New DConecta
    dbConec.AbreConexion
    psSql = "stp_sel_InteresesDevengadosDetXCta '" & psCtaCod & "','" & Format(pdFechaIni, "yyyymmdd") & "','" & Format(pdFechaFin, "yyyymmdd") & "'"
    Set InteresDevengadoDetXCta = dbConec.CargaRecordSet(psSql)
    Set dbConec = Nothing
    Exit Function
RepError:
    Call RaiseError(MyUnhandledError, "DAjusteCont:InteresDevengadoDetXCta Method")
End Function
Public Function InteresSuspensoDetXCta(ByVal psCtaCod As String, ByVal pdFechaIni As Date, ByVal pdFechaFin As Date) As ADODB.Recordset
On Error GoTo RepError
    Set dbConec = New DConecta
    dbConec.AbreConexion
    psSql = "stp_sel_InteresesSuspensoDetXCta '" & psCtaCod & "','" & Format(pdFechaIni, "yyyymmdd") & "','" & Format(pdFechaFin, "yyyymmdd") & "'"
    Set InteresSuspensoDetXCta = dbConec.CargaRecordSet(psSql)
    Set dbConec = Nothing
    Exit Function
RepError:
    Call RaiseError(MyUnhandledError, "DAjusteCont:InteresSuspensoDetXCta Method")
End Function
Public Function ProvisionCarteraDetXCta(ByVal psCtaCod As String, ByVal pdFechaIni As Date, ByVal pdFechaFin As Date) As ADODB.Recordset
On Error GoTo RepError
    Set dbConec = New DConecta
    dbConec.AbreConexion
    psSql = "stp_sel_ProvisionCarteraDetXCta '" & psCtaCod & "','" & Format(pdFechaIni, "yyyymmdd") & "','" & Format(pdFechaFin, "yyyymmdd") & "'"
    Set ProvisionCarteraDetXCta = dbConec.CargaRecordSet(psSql)
    Set dbConec = Nothing
    Exit Function
RepError:
    Call RaiseError(MyUnhandledError, "DAjusteCont:ProvisionCarteraDetXCta Method")
End Function
'END PASI
'PASI20170424 **********************
Public Function CargaAjusteDeContabilidadxCredito(pnTipoAsiento As Integer, psFecha As Date, pnMoneda As Integer, pnTipCambio As Currency) As Recordset
    On Error GoTo ErrCargaAjusteDeContabilidadxCredito
        Set dbConec = New DConecta
        dbConec.AbreConexion
        Select Case pnTipoAsiento
            Case 28
                psSql = "stp_sel_CargaAjusteRevIntDevxZonaInundada '" & pnMoneda & "','" & Format(psFecha, "YYYY/MM/DD") & "'," & pnTipCambio
        End Select
        Set CargaAjusteDeContabilidadxCredito = dbConec.CargaRecordSet(psSql)
        Set dbConec = Nothing
    Exit Function
ErrCargaAjusteDeContabilidadxCredito:
    Call RaiseError(MyUnhandledError, "DAjusteCont:CargaAjusteDeContabilidadxCredito Method")
End Function
Public Function RegistrarColocAjusteDeContabilidad(ByVal psUser As String, ByVal pnMovNro As Long) As Long
    On Error GoTo ErrRegistreColocAjusteDeContabilidad
    Dim rs As ADODB.Recordset
    Set dbConec = New DConecta
    dbConec.AbreConexion
    psSql = "stp_ins_RegistraColocAjusteDeContabilidad '" & psUser & "'," & pnMovNro
    Set rs = dbConec.CargaRecordSet(psSql)
    RegistrarColocAjusteDeContabilidad = IIf(Not (rs.EOF And rs.BOF), rs!nId, 0)
    Set dbConec = Nothing
    Exit Function
ErrRegistreColocAjusteDeContabilidad:
    Call RaiseError(MyUnhandledError, "DAjusteCont:RegistreColocAjusteDeContabilidad Method")
End Function
Public Function RegistrarColocAjusteDeContabilidadDet(ByVal pnId As Integer, ByVal psCtaCod As String, ByVal psCtaContDebe As String, ByVal psCtaContHaber As String, ByVal pnImporte As Currency)
    On Error GoTo ErrRegistreColocAjusteDeContabilidadDet
    Set dbConec = New DConecta
    dbConec.AbreConexion
    psSql = "stp_ins_RegistraColocAjusteDeContabilidadDet " & pnId & ",'" & psCtaCod & "','" & psCtaContDebe & "','" & psCtaContHaber & "'," & pnImporte
    dbConec.Ejecutar psSql
    Set dbConec = Nothing
    Exit Function
ErrRegistreColocAjusteDeContabilidadDet:
    Call RaiseError(MyUnhandledError, "DAjusteCont:RegistrarColocAjusteDeContabilidadDet Method")
End Function
'PASI END***

Public Function CargaAjusteReporteDet(ByVal psFecha As String, ByVal pnTipoAsiento As Integer, ByVal psMonOpeTpo As String, Optional psSubProc As String = "") As Recordset 'NAGL 202008 Cambió de psTipo a pnTipoAsiento 'NAGL202008 Agregó psOtrosSubProc
On Error GoTo CargaAjusteReporteDetErr
Set dbConec = New DConecta
Dim psRep As String
dbConec.AbreConexion
psRep = "Rep"
If pnTipoAsiento = 31 Then 'Reprog
  psSql = "exec stp_sel_ReporteCierreCarteraReprog '" & Format(psFecha, "MM/dd/YYYY") & "','" & psMonOpeTpo & "'"
ElseIf pnTipoAsiento = 32 And psSubProc = "" Then 'DifCancel
  psSql = "exec stp_ins_ObtieneDiferidosCredCancel '" & Format(psFecha, "MM/dd/YYYY") & "','" & psMonOpeTpo & "','" & psRep & "'"
ElseIf pnTipoAsiento = 32 And psSubProc = "SNeg" Then 'Dif Saldos Negativos
  psSql = "exec stp_ins_DiferidosSaldosNegativos '" & Format(psFecha, "MM/dd/YYYY") & "','" & psMonOpeTpo & "','" & psRep & "'"
ElseIf pnTipoAsiento = 33 Then
  psSql = "exec stp_ins_ObtieneNivelacionCuentasOrden_RCD '" & Format(psFecha, "MM/dd/YYYY") & "','" & psMonOpeTpo & "','" & psRep & "'"
ElseIf pnTipoAsiento = 34 Then 'NAGL 202102
  psSql = "exec stp_sel_ReporteCierreProvReprogCOVID '" & Format(psFecha, "MM/dd/YYYY") & "','" & psMonOpeTpo & "'" 'NAGL 202102 Según Acta N°017-2021
End If 'NAGL 202008 Según ACTA N°063-2020

Set CargaAjusteReporteDet = dbConec.CargaRecordSet(psSql)
Set dbConec = Nothing
Exit Function
CargaAjusteReporteDetErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:CargaAjusteReporteDet Method")
End Function 'NAGL 202007 Según ACTA N°046-2020

Public Sub InsertaInteresesDiferidos(ByVal psFecha As String, ByVal psOpecod As String)
On Error GoTo InsertaInteresesDiferidosErr
Set dbConec = New DConecta
    'Procede a Insertar inf. en ReporteInteresDiferidoAmpliados con psFecha
    dbConec.AbreConexion
    psSql = "Exec stp_ins_IntDifAmpliados '" & Format(psFecha, "MM/dd/YYYY") & "','" & psOpecod & "'"
    dbConec.Ejecutar psSql
    dbConec.CierraConexion
    Set dbConec = Nothing
Exit Sub
InsertaInteresesDiferidosErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:InsertaInteresesDiferidos Method")
End Sub 'NAGL 202008 Según Acta N°063-2020

Public Function CargaAsientosOtrosProcesos(ByVal psFecha As String, ByVal psOpecod As String, ByVal psMovNroOrigen As String, pnTipoAsiento As Integer) As String
On Error GoTo CargaAsientosOtrosProcesosErr
Set dbConec = New DConecta
Dim oImp As New NContImprimir
Dim rs As New ADODB.Recordset
Dim psRep As String, psObserv As String, sImpre As String, psMovNro As String
Dim pnMovNro As Long
dbConec.AbreConexion
psMovNro = ""
sImpre = ""
If pnTipoAsiento = 32 Then
    'INSERTA MOVIMIENTO
    '--DIF CRED CANCEL
    psRep = "Asi"
    psSql = "exec stp_ins_ObtieneDiferidosCredCancel '" & Format(psFecha, "MM/dd/YYYY") & "','" & psOpecod & "','" & psRep & "','" & psMovNroOrigen & "'"
    Set rs = dbConec.CargaRecordSet(psSql)
    psObserv = rs!cObservacion
    pnMovNro = rs!nMovNro
    
    If psObserv = "" Then
        '--DIF SALDOS NEGATIVOS
        Set rs = Nothing
        psSql = "exec stp_ins_DiferidosSaldosNegativos '" & Format(psFecha, "MM/dd/YYYY") & "','" & psOpecod & "','" & psRep & "'," & pnMovNro & ",'" & psMovNroOrigen & "'"
        Set rs = dbConec.CargaRecordSet(psSql)
        psObserv = rs!cObservacion
        pnMovNro = rs!nMovNro
    End If
    
ElseIf pnTipoAsiento = 33 Then
    'INSERTA MOVIMIENTO
    '--NIVELACION CUENTAS ORDEN - RCD
    psRep = "Asi"
    psSql = "exec stp_ins_ObtieneNivelacionCuentasOrden_RCD '" & Format(psFecha, "MM/dd/YYYY") & "','" & psOpecod & "','" & psRep & "','" & psMovNroOrigen & "'"
    Set rs = dbConec.CargaRecordSet(psSql)
    psObserv = rs!cObservacion
    pnMovNro = rs!nMovNro
End If

    If pnMovNro > 0 Then
        'INSERTA CUENTAS - PROASIENTO
        Set rs = Nothing
        psSql = "exec stp_ins_IngresarAsientoDNxnMovNro '" & Format(psFecha, "MM/dd/YYYY") & "'," & pnMovNro & ",'" & psOpecod & "'"
        Set rs = dbConec.CargaRecordSet(psSql)
        psObserv = rs!cObservacion
        pnMovNro = rs!nMovNro
        If pnMovNro > 0 Then
            'GENERA ASIENTOS
            Set rs = Nothing
            psSql = "exec stp_ins_InsertaMovCta_MovMExnMovNroxAsientoDN '" & Format(psFecha, "MM/dd/YYYY") & "'," & pnMovNro & ""
            Set rs = dbConec.CargaRecordSet(psSql)
            psObserv = rs!cObservacion
            psMovNro = rs!cMovNro
            If psObserv <> "" Then
                MsgBox psObserv, vbInformation, "Aviso"
                Exit Function
            End If
        Else
           If IIf(IsNull(rs!cCtaCnt), "", rs!cCtaCnt) <> "" Then
                sImpre = oImp.ImprimeCuentasObservacion(rs, pnTipoAsiento)
                MsgBox "Observación Contable..!! Seleccione Aceptar, para ver el detalle.", vbInformation, "Aviso"
           Else
                MsgBox psObserv, vbInformation, "Aviso"
                Exit Function
           End If
        End If
    Else
       MsgBox psObserv, vbInformation, "Aviso"
       Exit Function
    End If
    
 If sImpre = "" Then
    CargaAsientosOtrosProcesos = psMovNro
 Else
    CargaAsientosOtrosProcesos = sImpre
 End If
 Set rs = Nothing
 Set dbConec = Nothing
 Exit Function
CargaAsientosOtrosProcesosErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:CargaAsientosOtrosProcesos Method")
End Function 'NAGL 202008 Según Acta N°063-2020

Public Function CargaListaFechasColReprogCOVID(ByVal psFecha As String, ByVal psTipo As String) As Recordset
On Error GoTo CargaListaFechasColReprogCOVIDErr
Set dbConec = New DConecta
dbConec.AbreConexion
    psSql = "exec stp_sel_ListadoFechasColumnsReprogCOVID '" & Format(psFecha, "MM/dd/YYYY") & "','" & psTipo & "'"
Set CargaListaFechasColReprogCOVID = dbConec.CargaRecordSet(psSql)
Set dbConec = Nothing
Exit Function
CargaListaFechasColReprogCOVIDErr:
   Call RaiseError(MyUnhandledError, "DAjusteCont:CargaListaFechasColReprogCOVID Method")
End Function 'NAGL 202102 Según ACTA N°017-2021
