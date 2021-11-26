VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCajero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsConexion As String
Dim vsBaseComunes As String
Dim vsBasePesonas As String

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As DClassIni
Set oIni = New DClassIni
vsConexion = oIni.CadenaConexion
vsBaseComunes = oIni.BaseComunes
vsBasePesonas = oIni.BasePersonas
Set oIni = Nothing
End Sub

Public Function GetLibroCajaDatos(pdFecha As Date, pnMoneda As Moneda, pbBilletes As Boolean, psAgencia As String, pgbBitCental As Boolean) As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If pgbBitCental Then
        sql = " Select Sum(IsNull(nMonto,0)) Monto, nEfectivoValor From MovUserEfectivo a" _
            & " Right Join Efectivo b On a.cEfectivoCod = b.cEfectivoCod" _
            & " Inner Join Mov c On a.nMovNro = c.nMovNro" _
            & " Where c.cMovNro like '" & Format(pdFecha, gsFormatoMovFecha) & "%' And c.nMovFlag = " & MovFlag.gMovFlagVigente & " And b.cEfectivoCod Like '" & Trim(Str(pnMoneda)) & IIf(pbBilletes, "B", "M") & "%' And c.cOpeCod in ('901016','901007') And Substring(c.cMovNro,18,2) = '" & psAgencia & "' " _
            & " group by b.nEfectivoValor Order by nEfectivoValor Desc "
        oCon.AbreConexion
        
    Else
        sql = " Select nMoneda nEfectivoValor, Sum (nMoneda * nCantidad) Monto  from billetaje where datediff(day,dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 and cTipMoneda = '" & IIf(pbBilletes, "B", "M") & "'" _
            & " And cMoneda = '" & Trim(Str(pnMoneda)) & "' group by nMoneda order by nMoneda desc"

        'oCon.AbreConexionRemota psAgencia
    End If
    
    Set GetLibroCajaDatos = oCon.CargaRecordSet(sql)
End Function

Public Function GetValorBilletaje(pdFecha As Date, pnMoneda As Moneda, psAgencia As String, pgbBitCental As Boolean, pbAyer As Boolean, Optional pbHistoCtaContb As Boolean = False) As Currency 'JACA 20111229 se agregó pbHistoCtaContb
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If pgbBitCental Then
        oCon.AbreConexion
        
        sql = " Select IsNull(Sum(nMonto),0) Monto From MovUserEfectivo a" _
            & " Inner Join Mov c On a.nMovNro = c.nMovNro" _
            & " Where left(c.cMovNro,8) = (Select max(left(c.cMovNro,8))" _
            & " From MovUserEfectivo a" _
            & " Inner Join Mov c On a.nMovNro = c.nMovNro" _
            & "  Where c.cMovNro like  '" & Format(pdFecha, gsFormatoMovFecha) & "%' And c.nMovFlag = 0 And a.cEfectivoCod Like '1%' And c.cOpeCod in ('901016','901007')  And Substring(c.cMovNro,18,2) = '" & psAgencia & "' " _
            & " ) And c.nMovFlag = " & MovFlag.gMovFlagVigente & " And a.cEfectivoCod Like '" & Trim(Str(pnMoneda)) & "%' And c.cOpeCod in ('901016','901007')  And Substring(c.cMovNro,18,2) = '" & psAgencia & "'"
    Else
        'oCon.AbreConexionRemota psAgencia
        sql = " Select  IsNull(Sum (nMoneda * nCantidad),0) Monto from billetaje where datediff(day,dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 And cMoneda = '" & Trim(Str(pnMoneda)) & "'"
    End If
    
    Set rs = oCon.CargaRecordSet(sql)
    
    If pbAyer Then
        If rs.Fields(0) = 0 Then
            'JACA 20111229***************
            
            'PASI20150114
            Dim CtaNue As String
            CtaNue = ObtienePuenteAgexLibroCaja("11" & Trim(Str(pnMoneda)) & "102" & psAgencia)
            'end PASI
            If Not pbHistoCtaContb Then
                'Modificado PASI20150114
                'sql = " Select dbo.getsaldocta('" & Format(pdFecha, gsFormatoFecha) & "','11" & Trim(Str(pnMoneda)) & "102" & psAgencia & "'," & Trim(Str(pnMoneda)) & ") Monto "
                sql = " Select dbo.getsaldocta('" & Format(pdFecha, gsFormatoFecha) & "'," & IIf(CtaNue = "00", "'11" & Trim(Str(pnMoneda)) & "102" & psAgencia & "'", "'" & CtaNue & "'") & "," & Trim(Str(pnMoneda)) & ") Monto "
                'end PASI
            Else
                'Arreglar cuando se cambia de plan contable
                'sql = " Select dbo.Histo_getsaldocta('" & Format(pdFecha, gsFormatoFecha) & "','11" & Trim(Str(pnMoneda)) & "102" & psAgencia & "'," & Trim(Str(pnMoneda)) & ") Monto "
                
                'Modificado PASI20150114
                'sql = " Select dbo.getsaldocta('" & Format(pdFecha, gsFormatoFecha) & "','11" & Trim(Str(pnMoneda)) & "102" & psAgencia & "'," & Trim(Str(pnMoneda)) & ") Monto "
                sql = " Select dbo.getsaldocta('" & Format(pdFecha, gsFormatoFecha) & "'," & IIf(CtaNue = "00", "'11" & Trim(Str(pnMoneda)) & "102" & psAgencia & "'", "'" & CtaNue & "'") & "," & Trim(Str(pnMoneda)) & ") Monto "
                'end PASI
            End If
            'JACA END**********************
            Set rs = oCon.CargaRecordSet(sql)
        End If
    End If
    
    GetValorBilletaje = rs.Fields(0)
End Function

Public Function GetValorMovimiento(pdFecha As Date, pbIngreso As Boolean, pnMoneda As Moneda, psAgencia As String, pgbBitCental As Boolean) As Currency
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim lsIngEgr As String
    Dim lsCodOpeCV As String
    
    If pgbBitCental Then
        If pbIngreso Then
            lsIngEgr = "I"
            If pnMoneda = gMonedaNacional Then
                lsCodOpeCV = "('900023','900026','900029') "
            Else
                lsCodOpeCV = "('900022','900025','900028')"
            End If
        Else
            lsIngEgr = "E"
            If pnMoneda = gMonedaNacional Then
                lsCodOpeCV = "('900022','900025','900028')"
            Else
                lsCodOpeCV = "('900023','900026','900029') "
            End If
        End If
        
        
        oCon.AbreConexion
        sql = " Select Sum(nMonto) Monto From mov a " _
            & " Inner Join MovCol b on a.nMovNro = b.nMovNro" _
            & " Left Join MovDoc c on a.nMovNro = c.nMovNro And nDocTpo = 47" _
            & " Where c.nMovNro is null and a.cMovNro like '" & Format(pdFecha, gsFormatoMovFecha) & "%' and b.cOpeCod in" _
            & " (Select cOpeCod from grupoope a" _
            & "     inner join gruposope b on a.cGrupoCod = b.cGrupoCod" _
            & "     where nEfectivo = 1 And cIngEgr = '" & lsIngEgr & "')" _
            & " And Substring(b.cCtaCod,9,1) = '" & Trim(Str(pnMoneda)) & "' And Substring(a.cMovNro,18,2) = '" & psAgencia & "' And nmovflag = 0 "
        
        sql = sql & " UNION ALL Select Sum(nMonto) Monto from mov a" _
            & " Inner Join MovCap b on a.nMovNro = b.nMovNro" _
            & " Left Join MovDoc c on a.nMovNro = c.nMovNro And nDocTpo = 47" _
            & " Where c.nMovNro is null and a.cMovNro like '" & Format(pdFecha, gsFormatoMovFecha) & "%' and b.cOpeCod in" _
            & " (Select cOpeCod from grupoope a" _
            & "  inner join gruposope b on a.cGrupoCod = b.cGrupoCod" _
            & "  where nEfectivo = 1 And cIngEgr = '" & lsIngEgr & "')" _
            & " And Substring(b.cCtaCod,9,1) = '" & Trim(Str(pnMoneda)) & "' And Substring(a.cMovNro,18,2) = '" & psAgencia & "' And nmovflag = 0 "
    
        sql = sql & " UNION ALL  Select Sum(abs(nMovImporte)) Monto from mov a" _
            & " Inner Join movopevarias b on a.nMovNro = b.nMovNro" _
            & " Where a.cMovNro like '" & Format(pdFecha, gsFormatoMovFecha) & "%' and a.cOpeCod in" _
            & " (Select cOpeCod from grupoope a" _
            & " inner join gruposope b on a.cGrupoCod = b.cGrupoCod" _
            & " where nEfectivo = 1 And cIngEgr = '" & lsIngEgr & "')" _
            & " And b.nMoneda = '" & Trim(Str(pnMoneda)) & "'  And Substring(a.cMovNro,18,2) = '" & psAgencia & "' And nmovflag = 0 "
    
        sql = sql & " UNION ALL Select Sum(nMonto) Monto from mov a" _
            & " Inner Join movcmac b on a.nMovNro = b.nMovNro" _
            & " Where a.cMovNro like '" & Format(pdFecha, gsFormatoMovFecha) & "%' and a.cOpeCod in" _
            & " (Select cOpeCod from grupoope a" _
            & "     inner join gruposope b on a.cGrupoCod = b.cGrupoCod" _
            & "     where nEfectivo = 1 And cIngEgr = '" & lsIngEgr & "')" _
            & " And b.nMoneda = '" & Trim(Str(pnMoneda)) & "'  And Substring(a.cMovNro,18,2) = '" & psAgencia & "' And nmovflag = 0 "
    
        sql = sql & " UNION ALL Select isnull(sum(round(" & IIf(pnMoneda = gMonedaNacional, " nMovImporte * nMovTpoCambio ", " nMovImporte ") & ",2)),0) Monto from movcompraventa a" _
            & " inner join mov b on a.nmovnro = b.nmovnro" _
            & " inner join movtpocambio c on a.nmovnro = c.nmovnro" _
            & " where cmovnro like '" & Format(pdFecha, gsFormatoMovFecha) & "%' And Substring(b.cMovNro,18,2) = '" & psAgencia & "' And copecod in " & lsCodOpeCV & "" _
            & " And nmovflag = 0"
    
        sql = sql & " UNION ALL Select sum(nMonto) from movservicios a" _
            & " inner join mov b on a.nmovnro = b.nmovnro" _
            & " where cmovnro like '" & Format(pdFecha, gsFormatoMovFecha) & "%' And Substring(b.cMovNro,18,2) = '" & psAgencia & "' " _
            & " And nmovflag = 0 And a.nMoneda = " & Trim(Str(pnMoneda)) & " And  b.copecod in (Select cOpeCod from grupoope a " _
            & "     inner join gruposope b on a.cGrupoCod = b.cGrupoCod" _
            & "     where nEfectivo = 1 And cIngEgr = '" & lsIngEgr & "')"
        
        
        sql = sql & " UNION ALL SELECT SUM(abs(nMontoHabDev)) As Efectivo" _
             & " FROM GrupoOpe As G" _
             & " JOIN GruposOpe GO ON GO.cGrupoCod = G.cGrupoCod" _
             & " JOIN (SELECT cOpeCod,   CASE WHEN cAgeOrig = '" & psAgencia & "' THEN  Efectivo*-1 ELSE Efectivo END AS nMontoHabDev,   0 AS Cheque, 0 AS OrdenPago, nTotalMov" _
             & " FROM (SELECT M.cOpeCod, COUNT(*) AS nTotalMov, SUM(MH.nMovImporte) AS Efectivo, ISNULL(MH.cAgeOrig,'') AS cAgeOrig" _
             & " FROM Mov M JOIN MovHabilitacion MH ON MH.nMovNro = M.nMovNro" _
             & " JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod" _
             & " WHERE       M.nMovFlag = 0 And MH.nMoneda = " & Trim(Str(pnMoneda)) & "     AND (MH.cAgeOrig = '" & psAgencia & "' ) AND SUBSTRING(M.cMovNro,1,8) = '" & Format(pdFecha, gsFormatoMovFecha) & "'" _
             & " GROUP BY M.cOpeCod, MH.cAgeOrig) AS HabDev" _
             & " UNION SELECT cOpeCod,   CASE WHEN cAgeOrig = '" & psAgencia & "' THEN  Efectivo*-1 ELSE Efectivo END AS nMontoHabDev,   0 AS Cheque, 0 AS OrdenPago, nTotalMov" _
             & " FROM (SELECT M.cOpeCod, COUNT(*) AS nTotalMov, SUM(MH.nMovImporte) AS Efectivo,      ISNULL(MH.cAgeOrig,'') AS cAgeOrig" _
             & " FROM Mov M JOIN MovHabilitacion MH ON MH.nMovNro = M.nMovNro" _
             & " JOIN OpeTpo O ON O.cOpeCod = M.cOpeCod    WHERE       M.nMovFlag = 0 And MH.nMoneda = " & Trim(Str(pnMoneda)) & "      AND (MH.cAgeDest = '" & psAgencia & "')" _
             & " AND EXISTS (SELECT M.nMovNro FROM Mov M1" _
             & " JOIN MOVREF MR ON MR.nMovNro = M1.nMovNro" _
             & " WHERE SUBSTRING(M1.cMovNro,1,8) = '" & Format(pdFecha, gsFormatoMovFecha) & "' AND MR.nMovNroREF = M.nMovNro AND M1.nMovFlag = 0)" _
             & " GROUP BY M.cOpeCod, MH.cAgeOrig) AS HabDev ) As OpeHD ON OpeHD.cOpeCod = GO.cOpeCod" _
             & " Where G.cGrupoCod LIKE '06%'" _
             & " GROUP BY cGrupoNombre, GO.cGrupoCod having SUM(nMontoHabDev) " & IIf(pbIngreso, " > 0", " < 0")
        
        
        sql = " Select Sum(IsNull(Monto,0)) Monto From (" & sql & ")  As AAA"
    Else
        
        If pbIngreso Then
            lsIngEgr = "I"
            If pnMoneda = gMonedaNacional Then
                lsCodOpeCV = "230102"
            Else
                lsCodOpeCV = "230101"
            End If
        Else
            lsIngEgr = "E"
            If pnMoneda = gMonedaNacional Then
                lsCodOpeCV = "230101"
            Else
                lsCodOpeCV = "230102"
            End If
        End If
        
        
        'oCon.AbreConexionRemota psAgencia
        oCon.AbreConexion
        If pnMoneda = gMonedaNacional Then
            sql = " Select Sum(Monto) Monto From" _
                & " (Select 'A' c, IsNull(sum(abs(nmontran)),0) Monto from dbcomunes..opegru a" _
                & " inner join dbcomunes..operacion b on a.cGruProd = b.cGruProd" _
                & " inner join trandiariaconsol c on b.ccodope = c.ccodope and datediff(day,dfectran,'" & Format(pdFecha, gsFormatoFecha) & "') = 0" _
                & " Where cIngEgr = '" & lsIngEgr & "' And cMovCaj = 'S' And a.cGruProd Not In ('07001001','07002001') And Left(a.cGruProd,2) <> '02' " _
                & " and substring(ccodcta,6,1) = '" & Trim(Str(pnMoneda)) & "' and (cflag is null or cflag in ('1','2')) And b.ccodope not like '25%' And nTipTra in (1,3)" _
                & " Union" _
                & " Select 'B' c, IsNull(sum(round(abs(nmontran*nTipCambio),2)),0) Monto from trandiariaconsol where  ccodope = '" & lsCodOpeCV & "' and cflag is null" _
                & " and datediff(day,dfectran,'" & Format(pdFecha, gsFormatoFecha) & "') = 0) As AAA"
        
        Else
            sql = " Select Sum(Monto) Monto From" _
                & " (Select 'A' c, IsNull(sum(abs(nmontran)),0) Monto from dbcomunes..opegru a" _
                & " inner join dbcomunes..operacion b on a.cGruProd = b.cGruProd" _
                & " inner join trandiariaconsol c on b.ccodope = c.ccodope and datediff(day,dfectran,'" & Format(pdFecha, gsFormatoFecha) & "') = 0" _
                & " Where cIngEgr = '" & lsIngEgr & "' And cMovCaj = 'S' And a.cGruProd Not In ('07001001','07002001')  And Left(a.cGruProd,2) <> '02'" _
                & " and substring(ccodcta,6,1) =  '" & Trim(Str(pnMoneda)) & "' and (cflag is null or cflag in ('1','2')) And b.ccodope not like '25%'  And nTipTra in (1,3)" _
                & " Union" _
                & " Select 'B' c, IsNull(sum(abs(nmontran)),0) Monto from trandiariaconsol where  ccodope = '" & lsCodOpeCV & "' and cflag is null" _
                & " and datediff(day,dfectran,'" & Format(pdFecha, gsFormatoFecha) & "') = 0) As AAA"
        End If
    End If
    
    GetValorMovimiento = oCon.CargaRecordSet(sql).Fields(0)
End Function

Public Function GetHabDev(pdFecha As Date, pbIngreso As Boolean, pnMoneda As Moneda, psAgencia As String, pgbBitCental As Boolean) As Currency
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim lsOperacion As String
    
    If Not pgbBitCental Then
        'oCon.AbreConexionRemota psAgencia
        
        If pbIngreso Then
            If pnMoneda = gMonedaNacional Then
                lsOperacion = "250301"
            Else
                lsOperacion = "250302"
            End If
        Else
            If pnMoneda = gMonedaNacional Then
                lsOperacion = "250401"
            Else
                lsOperacion = "250402"
            End If
        End If
        
        sql = " Select IsNull(Sum(nMonTran),0) Monto From TransAho where cCodOpe = '" & lsOperacion & "' And Datediff(day, dFectran,'" & Format(pdFecha, gsFormatoFecha) & "') = 0  And cflag is null"
    Else
        oCon.AbreConexion
        If pbIngreso Then
            If pnMoneda = gMonedaNacional Then
                lsOperacion = "901002"
            Else
                lsOperacion = "901002"
            End If
        Else
            If pnMoneda = gMonedaNacional Then
                lsOperacion = "901003"
            Else
                lsOperacion = "901003"
            End If
        End If
        sql = " select isnull(sum(nMonto),0) from mov m join movuserefectivo me on me.nmovnro = m.nmovnro" _
            & " where cEfectivoCod LIKE '" & pnMoneda & "%' and copecod like '" & lsOperacion & "' and cmovnro like '" & Format(pdFecha, gsFormatoMovFecha) & "' and m.nMovFlag = 0 "
    End If

    GetHabDev = oCon.CargaRecordSet(sql).Fields(0)
End Function

Public Function GetSobFal(pdFecha As Date, pbIngreso As Boolean, pnMoneda As Moneda, psAgencia As String, pgbBitCental As Boolean) As Currency
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim lsOperacion As String
    
    If Not pgbBitCental Then
        'oCon.AbreConexionRemota psAgencia
        
        If pbIngreso Then
            lsOperacion = "260102"
        Else
            lsOperacion = "260103"
        End If
        
        sql = " Select IsNull(Sum(nMonTran),0) Monto From TransAho where cCodOpe = '" & lsOperacion & "' And Datediff(day, dFectran,'" & Format(pdFecha, gsFormatoFecha) & "') = 0 And Substring(cCodCta,6,1) = '" & Trim(Str(pnMoneda)) & "' And cflag is null"
    
    End If
    
    GetSobFal = oCon.CargaRecordSet(sql).Fields(0)
End Function
'PASI20150114
Public Function ObtienePuenteAgexLibroCaja(ByVal psCtaCont As String) As String
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    sql = "stp_sel_ObtienePuenteAgexLibroCaja '" & psCtaCont & "'"
    Set rs = oCon.CargaRecordSet(sql)
    If Not rs.BOF And Not rs.EOF Then
         ObtienePuenteAgexLibroCaja = CStr(rs!CtaNueva)
    Else
         ObtienePuenteAgexLibroCaja = "00"
    End If
    Exit Function
End Function
'END PASI
