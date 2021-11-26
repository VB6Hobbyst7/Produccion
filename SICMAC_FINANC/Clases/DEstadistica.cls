VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DEstadistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function RecupCreditoByFecha(ByVal pdFecha As String, ByVal pnTipoCambio As Double) As Recordset
    Dim oConec As DConecta
    Dim sSQL As String
    
    sSQL = "Select *"
    sSQL = sSQL & " From ("
    sSQL = sSQL & " Select isnull(Sum(CS.nSaldoCap),0) as nMonto,"
    sSQL = sSQL & " isnull(Sum(P.nTasaInteres)/Count(CS.cCtaCod),0) as TasaPromedio,"
    sSQL = sSQL & " 'Prendario' as cDescripcion,"
    sSQL = sSQL & " 'Soles' as cMoneda"
    sSQL = sSQL & " From ColocacSaldo CS"
    sSQL = sSQL & " Inner Join Producto P on P.cCtaCod=CS.cCtaCod"
    sSQL = sSQL & " Where CS.dFecha='" & pdFecha & "' and SubString(CS.cCtaCod,6,3) in('305') and SubString(CS.cCtacod,9,1)='1' and"
    sSQL = sSQL & " CS.nPrdEstado in (2101,2104,2106,2107)"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select isnull(Sum(CS.nSaldoCap),0) as nMonto,"
    sSQL = sSQL & " isnull(Sum(P.nTasaInteres)/Count(CS.cCtaCod),0) as TasaPromedio,"
    sSQL = sSQL & " 'Credito Pequeña Empresa' as cDescripcion,"
    sSQL = sSQL & " 'Soles' as cMoneda"
    sSQL = sSQL & " From ColocacSaldo CS"
    sSQL = sSQL & " Inner Join Producto P on P.cCtaCod=CS.cCtaCod"
    sSQL = sSQL & " Where CS.dFecha='" & pdFecha & "' and SubString(CS.cCtaCod,6,3) in('101','201') and SubString(CS.cCtacod,9,1)='1'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select isnull(Sum(CS.nSaldoCap*" & pnTipoCambio & "),0) as nMonto,"
    sSQL = sSQL & " isnull(Sum(P.nTasaInteres)/Count(CS.cCtaCod),0),"
    sSQL = sSQL & " 'Credito Pequeña Empresa' as cDescripcion,"
    sSQL = sSQL & " 'Dolares' as cMoneda"
    sSQL = sSQL & " From ColocacSaldo CS"
    sSQL = sSQL & " Inner Join Producto P on P.cCtaCod=CS.cCtaCod"
    sSQL = sSQL & " Where CS.dFecha='" & pdFecha & "' and SubString(CS.cCtaCod,6,3) in('101','201') and SubString(CS.cCtacod,9,1)='2'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select isnull(Sum(CS.nSaldoCap),0) as nMonto,"
    sSQL = sSQL & " isnull(Sum(P.nTasaInteres)/Count(CS.cCtaCod),0) as TasaPromedio,"
    sSQL = sSQL & " 'Credito Personal' as cDescripcion,"
    sSQL = sSQL & "  'Soles' as cMoneda"
    sSQL = sSQL & " From ColocacSaldo CS"
    sSQL = sSQL & " Inner Join Producto P on P.cCtaCod=CS.cCtaCod"
    sSQL = sSQL & " Where CS.dFecha='" & pdFecha & "' and SubString(CS.cCtaCod,6,3) in ('301','302','304','423') and SubString(CS.cCtacod,9,1)='1'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select isnull(Sum(CS.nSaldoCap*" & pnTipoCambio & "),0) as nMonto,"
    sSQL = sSQL & " isnull(Sum(P.nTasaInteres)/Count(CS.cCtaCod),0) as TasaPromedio,"
    sSQL = sSQL & " 'Credito Personal' as cDescripcion,"
    sSQL = sSQL & " 'Dolares' as cMoneda"
    sSQL = sSQL & " From ColocacSaldo CS"
    sSQL = sSQL & " Inner Join Producto P on P.cCtaCod=CS.cCtaCod"
    sSQL = sSQL & " Where CS.dFecha='" & pdFecha & "' and SubString(CS.cCtaCod,6,3) in ('301','302','304','423') and SubString(CS.cCtacod,9,1)='2'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select isnull(Sum(CS.nSaldoCap),0) as nMonto,"
    sSQL = sSQL & " isnull(Sum(P.nTasaInteres)/Count(CS.cCtaCod),0) as TasaPromedio,"
    sSQL = sSQL & " 'Credito Agricola' as cDescripcion,"
    sSQL = sSQL & " 'Soles' as cMoneda"
    sSQL = sSQL & " From ColocacSaldo CS"
    sSQL = sSQL & " Inner Join Producto P on P.cCtaCod=CS.cCtaCod"
    sSQL = sSQL & " Inner Join ColocacCred C on C.cCtaCod=CS.cCtaCod"
    sSQL = sSQL & " Where CS.dFecha='" & pdFecha & "' and SubString(CS.cCtaCod,6,3) in('102','202') and SubString(CS.cCtacod,9,1)='1' and"
    sSQL = sSQL & " (C.cRFA in ('RFA','RFC') Or C.cRFA iS NULL)"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select isnull(Sum(CS.nSaldoCap*" & pnTipoCambio & "),0) as nMonto,"
    sSQL = sSQL & " isnull(Sum(P.nTasaInteres)/Count(CS.cCtaCod),0),"
    sSQL = sSQL & " 'Credito Agricola' as cDescripcion,"
    sSQL = sSQL & " 'Dolares' as cMoneda"
    sSQL = sSQL & " From ColocacSaldo CS"
    sSQL = sSQL & " Inner Join Producto P on P.cCtaCod=CS.cCtaCod"
    sSQL = sSQL & " Inner Join ColocacCred C on C.cCtaCod=CS.cCtaCod"
    sSQL = sSQL & " Where CS.dFecha='" & pdFecha & "' and SubString(CS.cCtaCod,6,3) in('102','202') and SubString(CS.cCtacod,9,1)='2' and"
    sSQL = sSQL & " (C.cRFA in ('RFA','RFC') Or C.cRFA iS NULL)"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select isnull(Sum(CS.nSaldoCap),0) as nMonto,"
    sSQL = sSQL & " isnull(Sum(P.nTasaInteres)/Count(CS.cCtaCod),0) as TasaPromedio,"
    sSQL = sSQL & " 'Credito CTS' as cDescripcion,"
    sSQL = sSQL & " 'Soles' as cMoneda"
    sSQL = sSQL & " From ColocacSaldo CS"
    sSQL = sSQL & " Inner Join Producto P on P.cCtaCod=CS.cCtaCod"
    sSQL = sSQL & " Where CS.dFecha='" & pdFecha & "' and SubString(CS.cCtaCod,6,3) in('301') and SubString(CS.cCtacod,9,1)='1'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select isnull(Sum(CS.nSaldoCap*" & pnTipoCambio & "),0) as nMonto,"
    sSQL = sSQL & " isnull(Sum(P.nTasaInteres)/Count(CS.cCtaCod),0) as TasaPromedio,"
    sSQL = sSQL & " 'Credito CTS' as cDescripcion,"
    sSQL = sSQL & " 'Dolares' as cMoneda"
    sSQL = sSQL & " From ColocacSaldo CS"
    sSQL = sSQL & " Inner Join Producto P on P.cCtaCod=CS.cCtaCod"
    sSQL = sSQL & " Where CS.dFecha='" & pdFecha & "' and SubString(CS.cCtaCod,6,3) in('301') and SubString(CS.cCtacod,9,1)='2'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select isnull(Sum(CS.nSaldoCap),0) as nMonto,"
    sSQL = sSQL & " isnull(Sum(P.nTasaInteres)/Count(CS.cCtaCod),0) as TasaPromedio,"
    sSQL = sSQL & " 'Prestamos CMAC' as cDescripcion,"
    sSQL = sSQL & " 'Soles' as cMoneda"
    sSQL = sSQL & " From ColocacSaldo CS"
    sSQL = sSQL & " Inner Join Producto P on P.cCtaCod=CS.cCtaCod"
    sSQL = sSQL & " Where CS.dFecha='" & pdFecha & "' and SubString(CS.cCtaCod,6,3) in('320') and SubString(CS.cCtacod,9,1)='1'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select isnull(Sum(CS.nSaldoCap*" & pnTipoCambio & "),0) as nMonto,"
    sSQL = sSQL & " isnull(Sum(P.nTasaInteres)/Count(CS.cCtaCod),0) as TasaPromedio,"
    sSQL = sSQL & " 'Prestamos CMAC' as cDescripcion,"
    sSQL = sSQL & " 'Dolares' as cMoneda"
    sSQL = sSQL & " From ColocacSaldo CS"
    sSQL = sSQL & " Inner Join Producto P on P.cCtaCod=CS.cCtaCod"
    sSQL = sSQL & " Where CS.dFecha='" & pdFecha & "' and SubString(CS.cCtaCod,6,3) in('320') and SubString(CS.cCtacod,9,1)='2') X"
    sSQL = sSQL & " Order By cDescripcion,cMoneda"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set RecupCreditoByFecha = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function TipoCambio(ByVal pnMes As Integer) As Double
    Dim oConec As DConecta
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    sSQL = "Select nValFijo"
    sSQL = sSQL & " From TipoCambio"
    sSQL = sSQL & " Where dFecCamb=(Select  Max(dFecCamb)"
    sSQL = sSQL & " From TipoCambio"
    sSQL = sSQL & " Where Month(dFecCamb)=" & pnMes + 1 & ")"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        TipoCambio = rs!nValFijo
    End If
    Set rs = Nothing
End Function

Public Function GeneraRepCapEAFecha(ByVal pdFecha As String, ByVal pnTipoCambio As Double) As Recordset
    Dim sSQL As String
    Dim oCon As DConecta
    
    sFecha = Format(pdFecha, "YYYYMMdd")
    
sSQL = "SELECT NRO=1,REPORTE='AHORRO CORRIENTE',TOTALSOLES=ISNULL(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=1 THEN  CS.NSALDCNT ELSE 0 END ),0),"
sSQL = sSQL & " TOTALSOLESME=ISNULL(ROUND(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=2 THEN  CS.NSALDCNT ELSE 0 END )*" & pnTipoCambio & ",2),0) ,"
sSQL = sSQL & " TASAPROMSOL=ROUND(((isnull(SUM( ROUND( case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt end * (dbo.conviertetnaatea(ntasaint)/100),2)),0) / case when isnull(SUM( case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt  end ),0)>0 then SUM(case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt  end ) else 1  end   ) * 100),2),"
sSQL = sSQL & " TASAPROMDOL=ROUND(((isnull(SUM( ROUND( case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt end * (dbo.conviertetnaatea(ntasaint)/100),2)),0) / case when isnull(SUM( case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt  end ),0)>0 then SUM(case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt  end ) else 1  end    ) * 100),2)"
sSQL = sSQL & " FROM CAPSALDOSDIARIOS CS"
sSQL = sSQL & " JOIN CAPTACIONES C ON C.CCTACOD=CS.CCTACOD"
sSQL = sSQL & " LEFT JOIN (SELECT * FROM PRODUCTOPERSONA WHERE NPRDPERSRELAC=10 AND CPERSCOD='1080100907511') PE ON PE.CCTACOD=CS.CCTACOD"
sSQL = sSQL & " WHERE CONVERT(CHAR(8),DFECHA,112)='" & sFecha & "' AND PE.CCTACOD IS NULL"
sSQL = sSQL & " AND CS.CCTACOD LIKE '108__232%' AND C.NPERSONERIA NOT IN (4)"
sSQL = sSQL & " Union"
sSQL = sSQL & " SELECT NRO=2, REPORTE='PLAZO FIJO',TOTALSOLES=ISNULL(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=1 THEN  CS.NSALDCNT ELSE 0 END),0),"
sSQL = sSQL & " TOTALSOLESME=ISNULL(ROUND(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=2 THEN  CS.NSALDCNT ELSE 0 END )*" & pnTipoCambio & ",2),0) ,"
sSQL = sSQL & " TASAPROMSOL=ROUND(((isnull(SUM( ROUND( case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt end * (dbo.conviertetnaatea(ntasaint)/100),2)),0) / case when isnull(SUM( case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt  end ),0)>0 then SUM(case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt  end ) else 1  end   ) * 100),2),"
sSQL = sSQL & " TASAPROMDOL=ROUND(((isnull(SUM( ROUND( case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt end * (dbo.conviertetnaatea(ntasaint)/100),2)),0) / case when isnull(SUM( case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt  end ),0)>0 then SUM(case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt  end ) else 1  end    ) * 100),2)"
sSQL = sSQL & " FROM CAPSALDOSDIARIOS CS"
sSQL = sSQL & " JOIN CAPTACIONES C ON C.CCTACOD=CS.CCTACOD"
sSQL = sSQL & " WHERE CONVERT(CHAR(8),DFECHA,112)='" & sFecha & "'"
sSQL = sSQL & " AND CS.CCTACOD LIKE '108__233%' AND C.NPERSONERIA NOT IN (4)"
sSQL = sSQL & " Union"
sSQL = sSQL & " SELECT NRO=3,REPORTE='DEP. AHO. FONCODES', TOTALSOLES=ISNULL(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=1 THEN  CS.NSALDCNT ELSE 0 END),0),"
sSQL = sSQL & " TOTALSOLESME=ISNULL(ROUND(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=2 THEN  CS.NSALDCNT ELSE 0 END )*" & pnTipoCambio & ",2),0) ,"
sSQL = sSQL & " TASAPROMSOL=ROUND(((isnull(SUM( ROUND( case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt end * (dbo.conviertetnaatea(ntasaint)/100),2)),0) / case when isnull(SUM( case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt  end ),0)>0 then SUM(case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt  end ) else 1  end   ) * 100),2),"
sSQL = sSQL & " TASAPROMDOL=ROUND(((isnull(SUM( ROUND( case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt end * (dbo.conviertetnaatea(ntasaint)/100),2)),0) / case when isnull(SUM( case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt  end ),0)>0 then SUM(case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt  end ) else 1  end    ) * 100),2)"
sSQL = sSQL & " FROM CAPSALDOSDIARIOS CS"
sSQL = sSQL & " JOIN CAPTACIONES C ON C.CCTACOD=CS.CCTACOD"
sSQL = sSQL & " JOIN (SELECT * FROM PRODUCTOPERSONA WHERE NPRDPERSRELAC=10 AND CPERSCOD='1080100907511') PE ON PE.CCTACOD=CS.CCTACOD"
sSQL = sSQL & " WHERE CONVERT(CHAR(8),DFECHA,112)='" & sFecha & "'"
sSQL = sSQL & " AND CS.CCTACOD LIKE '108__232%' AND C.NPERSONERIA=2"
sSQL = sSQL & " Union"
sSQL = sSQL & " SELECT NRO=4,REPORTE='DEP. FOCMAN PF.', TOTALSOLES=isnull(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=1 THEN  CS.NSALDCNT ELSE 0 END),0),"
sSQL = sSQL & " TOTALSOLESME=ISNULL(ROUND(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=2 THEN  CS.NSALDCNT ELSE 0 END )*" & pnTipoCambio & ",2),0) ,"
sSQL = sSQL & " TASAPROMSOL=ROUND(((isnull(SUM( ROUND( case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt end * (dbo.conviertetnaatea(ntasaint)/100),2)),0) / case when isnull(SUM( case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt  end ),0)>0 then SUM(case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt  end ) else 1  end   ) * 100),2),"
sSQL = sSQL & " TASAPROMDOL=ROUND(((isnull(SUM( ROUND( case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt end * (dbo.conviertetnaatea(ntasaint)/100),2)),0) / case when isnull(SUM( case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt  end ),0)>0 then SUM(case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt  end ) else 1  end    ) * 100),2)"
sSQL = sSQL & " FROM CAPSALDOSDIARIOS CS"
sSQL = sSQL & " JOIN CAPTACIONES C ON C.CCTACOD=CS.CCTACOD"
sSQL = sSQL & " JOIN (SELECT * FROM PRODUCTOPERSONA WHERE NPRDPERSRELAC=10 AND CPERSCOD='1080100829587') PE ON PE.CCTACOD=CS.CCTACOD"
sSQL = sSQL & " WHERE CONVERT(CHAR(8),DFECHA,112)='" & sFecha & "'"
sSQL = sSQL & " Union"
sSQL = sSQL & " SELECT NRO=5,REPORTE='DEP. CMAC AHORRO',  TOTALSOLES=ISNULL(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=1 THEN  CS.NSALDCNT ELSE 0 END  ),0),"
sSQL = sSQL & " TOTALSOLESME=ISNULL(ROUND(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=2 THEN  CS.NSALDCNT ELSE 0 END)*" & pnTipoCambio & ",2),0) ,"
sSQL = sSQL & " TASAPROMSOL=ROUND(((isnull(SUM( ROUND( case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt end * (dbo.conviertetnaatea(ntasaint)/100),2)),0) / case when isnull(SUM( case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt  end ),0)>0 then SUM(case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt  end ) else 1  end   ) * 100),2),"
sSQL = sSQL & " TASAPROMDOL=ROUND(((isnull(SUM( ROUND( case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt end * (dbo.conviertetnaatea(ntasaint)/100),2)),0) / case when isnull(SUM( case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt  end ),0)>0 then SUM(case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt  end ) else 1  end    ) * 100),2)"
sSQL = sSQL & " FROM CAPSALDOSDIARIOS CS"
sSQL = sSQL & " JOIN CAPTACIONES C ON C.CCTACOD=CS.CCTACOD"
sSQL = sSQL & " WHERE CONVERT(CHAR(8),DFECHA,112)='" & sFecha & "'"
sSQL = sSQL & " AND CS.CCTACOD LIKE '108__232%' AND C.NPERSONERIA=4"
sSQL = sSQL & " Union"
sSQL = sSQL & " SELECT NRO=6, REPORTE='DEP. CMAC PLAZOFIJO',  TOTALSOLES=ISNULL(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=1 THEN  CS.NSALDCNT ELSE 0 END  ),0),"
sSQL = sSQL & " TOTALSOLESME=ISNULL(ROUND(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=2 THEN  CS.NSALDCNT ELSE 0 END )*" & pnTipoCambio & ",2),0) ,"
sSQL = sSQL & " TASAPROMSOL=ROUND(((isnull(SUM( ROUND( case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt end * (dbo.conviertetnaatea(ntasaint)/100),2)),0) / case when isnull(SUM( case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt  end ),0)>0 then SUM(case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt  end ) else 1  end   ) * 100),2),"
sSQL = sSQL & " TASAPROMDOL=ROUND(((isnull(SUM( ROUND( case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt end * (dbo.conviertetnaatea(ntasaint)/100),2)),0) / case when isnull(SUM( case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt  end ),0)>0 then SUM(case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt  end ) else 1  end    ) * 100),2)"
sSQL = sSQL & " FROM CAPSALDOSDIARIOS CS"
sSQL = sSQL & " JOIN CAPTACIONES C ON C.CCTACOD=CS.CCTACOD"
sSQL = sSQL & " WHERE CONVERT(CHAR(8),DFECHA,112)='" & sFecha & "'"
sSQL = sSQL & " AND CS.CCTACOD LIKE '108__233%' AND C.NPERSONERIA=4"
sSQL = sSQL & " Union"
sSQL = sSQL & " SELECT NRO=7, REPORTE='DEP. CTS',  TOTALSOLES=ISNULL(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=1 THEN  CS.NSALDCNT ELSE 0 END  ),0),"
sSQL = sSQL & " TOTALSOLESME=ISNULL(ROUND(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=2 THEN  CS.NSALDCNT ELSE 0 END )*" & pnTipoCambio & ",2),0) ,"
sSQL = sSQL & " TASAPROMSOL=ROUND(((isnull(SUM( ROUND( case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt end * (dbo.conviertetnaatea(ntasaint)/100),2)),0) / case when isnull(SUM( case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt  end ),0)>0 then SUM(case WHEN SUBSTRING(CS.CCTACOD,9,1)=1 then nSaldcnt  end ) else 1  end   ) * 100),2),"
sSQL = sSQL & " TASAPROMDOL=ROUND(((isnull(SUM( ROUND( case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt end * (dbo.conviertetnaatea(ntasaint)/100),2)),0) / case when isnull(SUM( case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt  end ),0)>0 then SUM(case WHEN SUBSTRING(CS.CCTACOD,9,1)=2 then nSaldcnt  end ) else 1  end    ) * 100),2)"
sSQL = sSQL & " FROM CAPSALDOSDIARIOS CS"
sSQL = sSQL & " JOIN CAPTACIONES C ON C.CCTACOD=CS.CCTACOD"
sSQL = sSQL & " WHERE CONVERT(CHAR(8),DFECHA,112)='" & sFecha & "'"
sSQL = sSQL & " AND CS.CCTACOD LIKE '108__234%'"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set GeneraRepCapEAFecha = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    

End Function


Public Function GeneraRepCapEAFechaConsol(ByVal pdFecha As String, ByVal pnTipoCambio As Double) As Recordset

    Dim sSQL As String
    Dim oConec As DConecta
    Dim sFecha As String

    sFecha = Format(pdFecha, "YYYYMMdd")
                    
    sSQL = "SELECT NRO=1,REPORTE='DEPOSITOS DEL PUBLICO',MONTO=ISNULL(SUM( CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=1 THEN  CS.NSALDCNT ELSE 0 END  ),0) + ISNULL( ROUND( SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=2 THEN  CS.NSALDCNT ELSE 0 END )*" & pnTipoCambio & ",2) ,0)"
    sSQL = sSQL & " FROM CAPSALDOSDIARIOS CS"
    sSQL = sSQL & " LEFT JOIN ( SELECT PP.* FROM PRODUCTOPERSONA PP JOIN PERSONA E ON E.CPERSCOD=PP.CPERSCOD"
    sSQL = sSQL & " WHERE (E.CPERSnombre like  'MUNICIPALIDAD%' )  AND CCTACOD LIKE '108__23[234]%' AND NPRDPERSRELAC=10  ) M ON M.CCTACOD=CS.CCTACOD"
    sSQL = sSQL & " JOIN CAPTACIONES C ON C.CCTACOD=CS.CCTACOD"
    sSQL = sSQL & " WHERE CONVERT(CHAR(8),DFECHA,112)='" & sFecha & "' AND M.CCTACOD IS NULL"
    sSQL = sSQL & " AND C.NPERSONERIA<>4"
    sSQL = sSQL & " AND CS.CCTACOD LIKE '108__23[234]%'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " SELECT NRO=2,REPORTE='DEPOSITOS TOTAL CONCEJO PROVINCIAL',MONTO=ISNULL(SUM( CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=1 THEN  CS.NSALDCNT ELSE 0 END  ),0) + ISNULL( ROUND( SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=2 THEN  CS.NSALDCNT ELSE 0 END )*" & pnTipoCambio & ",2) ,0)"
    sSQL = sSQL & " FROM CAPSALDOSDIARIOS CS"
    sSQL = sSQL & " JOIN ( SELECT PP.* FROM PRODUCTOPERSONA PP  JOIN PERSONA E ON E.CPERSCOD=PP.CPERSCOD"
    sSQL = sSQL & " WHERE (E.CPERSnombre like  'MUNICIPALIDAD%')  AND CCTACOD LIKE '108__23[234]%'"
    sSQL = sSQL & " AND NPRDPERSRELAC=10   ) M ON M.CCTACOD=CS.CCTACOD"
    sSQL = sSQL & " JOIN CAPTACIONES C ON C.CCTACOD=CS.CCTACOD"
    sSQL = sSQL & " WHERE CONVERT(CHAR(8),DFECHA,112)='" & sFecha & "'"
    sSQL = sSQL & " AND CS.CCTACOD LIKE '108__23[234]%'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " SELECT NRO=3,REPORTE='DEPOSITOS DISTRITO EMP.MUN.REG',MONTO=0"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " SELECT NRO=4,REPORTE='DEPOSITOS DE CMACS',MONTO=ISNULL(SUM( CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=1 THEN  CS.NSALDCNT ELSE 0 END  ),0) + ISNULL( ROUND(SUM(CASE WHEN SUBSTRING(CS.CCTACOD,9,1)=2 THEN  CS.NSALDCNT ELSE 0 END )*" & pnTipoCambio & ",2) ,0)"
    sSQL = sSQL & " FROM CAPSALDOSDIARIOS CS"
    sSQL = sSQL & " JOIN CAPTACIONES C ON C.CCTACOD=CS.CCTACOD"
    sSQL = sSQL & " WHERE CONVERT(CHAR(8),DFECHA,112)='" & sFecha & "'"
    sSQL = sSQL & " AND C.NPERSONERIA=4"
    sSQL = sSQL & " AND CS.CCTACOD LIKE '108__23[234]%'"
    sSQL = sSQL & " ORDER BY NRO"
    Set oConec = New DConecta
    oConec.AbreConexion
    Set GeneraRepCapEAFechaConsol = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
End Function
Public Function RecupConsolidado(ByVal pdFecha As String, ByVal pnTipoCambio As Double) As Recordset
    Dim oConec As DConecta
    Dim sSQL As String
    Dim sFechaFinal As String
    
    sFechaFinal = Mid(pdFecha, 7, 4) & Mid(pdFecha, 1, 2) & Mid(pdFecha, 4, 2)
    
    sSQL = "Select *"
    sSQL = sSQL & " From ("
    sSQL = sSQL & " Select Count(MD.cCtaCod) as nValor,' Desembolsos del Año' as cDescripcion,'Pymes' as cProducto"
    sSQL = sSQL & " From Mov M"
    sSQL = sSQL & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
    sSQL = sSQL & " Where (Left(M.cMovNro,8) between '" & Mid(pdFecha, 7, 4) & "0101' and '" & sFechaFinal & "') and SubString(MD.cCtaCod,6,3) in ('101','201') and"
    sSQL = sSQL & " MD.cOpeCod in ('100101','100102','100103','100104','100105') and MD.nPrdConceptoCod=1000 and M.nMovFlag=0"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select Count(cCtaCod) as nValor,' Prestamos Vigentes' as cDescripcion,'Pymes' as cProducto"
    sSQL = sSQL & " From ColocacSaldo"
    sSQL = sSQL & " Where SubString(cCtaCod,6,3) in ('101','201') and dfecha='" & pdFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select Sum(Case When SubString(MD.cCtaCod,9,1)='1' Then MD.nMonto Else MD.nMonto*" & pnTipoCambio & " end) as nValor,' Monto Acumulado' as cDescripcion,'Pymes' as cProducto"
    sSQL = sSQL & " From Mov M"
    sSQL = sSQL & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
    sSQL = sSQL & " Where (Left(M.cMovNro,8) between '" & Mid(pdFecha, 7, 4) & "0101' and '" & sFechaFinal & "') and SubString(MD.cCtaCod,6,3) in ('101','201') and"
    sSQL = sSQL & " MD.cOpeCod in ('100101','100102','100103','100104','100105') and MD.nPrdConceptoCod=1000 and M.nMovFlag=0"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select Count(MD.cCtaCod) as nValor,' Desembolsos del Año' as cDescripcion,'Agricolas' as cProducto"
    sSQL = sSQL & " From Mov M"
    sSQL = sSQL & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
    sSQL = sSQL & " Where (Left(M.cMovNro,8) between '" & Mid(pdFecha, 7, 4) & "0101' and '" & sFechaFinal & "') and SubString(MD.cCtaCod,6,3) in ('102','202') and"
    sSQL = sSQL & " MD.cOpeCod in ('100101','100102','100103','100104','100105') and MD.nPrdConceptoCod=1000 and M.nMovFlag=0"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select Count(CS.cCtaCod) as nValor,' Prestamo Vigentes' as cDescripcion,'Agricolas' as cProducto"
    sSQL = sSQL & " From ColocacSaldo CS"
    sSQL = sSQL & " Inner Join ColocacCred C on C.cCtaCod=CS.cCtaCod"
    sSQL = sSQL & " Where SubString(CS.cCtaCod,6,3) in ('102','202') and CS.dfecha='" & pdFecha & "' and"
    sSQL = sSQL & " C.CRFA NOT IN ('RFA')"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select Sum(Case When SubString(MD.cCtaCod,9,1)='1' Then MD.nMonto Else MD.nMonto*" & pnTipoCambio & " end) as nValor,' Monto Acumulado' as cDescripcion,'Agricolas' as cProducto"
    sSQL = sSQL & " From Mov M"
    sSQL = sSQL & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
    sSQL = sSQL & " Where (Left(M.cMovNro,8) between '" & Mid(pdFecha, 7, 4) & "0101' and '" & sFechaFinal & "')  and SubString(MD.cCtaCod,6,3) in ('102','202') and"
    sSQL = sSQL & " MD.cOpeCod in ('100101','100102','100103','100104','100105') and MD.nPrdConceptoCod=1000 and M.nMovFlag=0"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select Count(MD.cCtaCod) as nValor,' Desembolsos del Año' as cDescripcion,'Prendario' as cProducto"
    sSQL = sSQL & " From Mov M"
    sSQL = sSQL & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
    sSQL = sSQL & " Where (Left(M.cMovNro,8) between '" & Mid(pdFecha, 7, 4) & "0101' and '" & sFechaFinal & "') and SubString(MD.cCtaCod,6,3) in ('305') and"
    sSQL = sSQL & " MD.cOpeCod in ('120201') and M.nMovFlag=0 and MD.nPrdConceptoCod =2100"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select Count(cCtaCod) as nValor,' Prestamos Vigentes' as cDescripcion,'Prendario' as cProducto"
    sSQL = sSQL & " From ColocacSaldo"
    sSQL = sSQL & " Where SubString(cCtaCod,6,3) in ('305') and dfecha='" & pdFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select Sum(Case When SubString(MD.cCtaCod,9,1)='1' Then MD.nMonto Else MD.nMonto*" & pnTipoCambio & " end) as nValor,' Monto Acumulado' as cDescripcion,'Prendario' as cProducto"
    sSQL = sSQL & " From Mov M"
    sSQL = sSQL & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
    sSQL = sSQL & " Where (Left(M.cMovNro,8) between '" & Mid(pdFecha, 7, 4) & "0101' and '" & sFechaFinal & "') and SubString(MD.cCtaCod,6,3) in ('305') and"
    sSQL = sSQL & " MD.cOpeCod in ('120201') and MD.nPrdConceptoCod=2100 and M.nMovFlag=0"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select Count(MD.cCtaCod) as nValor,' Desembolsos del Año' as cDescripcion,'Consumo' as cProducto"
    sSQL = sSQL & " From Mov M"
    sSQL = sSQL & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
    sSQL = sSQL & " Where (Left(M.cMovNro,8) between '" & Mid(pdFecha, 7, 4) & "0101' and '" & sFechaFinal & "') and SubString(MD.cCtaCod,6,3) Like '3%' and"
    sSQL = sSQL & " MD.cOpeCod in ('100101','100102','100103','100104','100105') and MD.nPrdConceptoCod=1000 and M.nMovFlag=0"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select Count(cCtaCod) as nValor,' Prestamos Vigentes' as cDescripcion,'Consumo' as cProducto"
    sSQL = sSQL & " From ColocacSaldo"
    sSQL = sSQL & " Where SubString(cCtaCod,6,3) Like '3%' and dfecha='" & pdFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " Select Sum(Case When SubString(MD.cCtaCod,9,1)='1' Then MD.nMonto Else MD.nMonto*" & pnTipoCambio & " end) as nValor,' Monto Acumulado' as cDescripcion,'Consumo' as cProducto"
    sSQL = sSQL & " From Mov M"
    sSQL = sSQL & " Inner Join MovColDet MD on M.nMovNro=MD.nMovNro"
    sSQL = sSQL & " Where (Left(M.cMovNro,8) between '" & Mid(pdFecha, 7, 4) & "0101' and '" & sFechaFinal & "') and SubString(MD.cCtaCod,6,3) Like '3%' and"
    sSQL = sSQL & " MD.cOpeCod in ('100101','100102','100103','100104','100105') and MD.nPrdConceptoCod=1000 and M.nMovFlag=0) X"
    sSQL = sSQL & " Order By cProducto"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set RecupConsolidado = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Public Function Recup_Caja(ByVal pdFecha As Date, ByVal pnTipoCambio As Double) As Recordset

    Dim sSQL As String
    Dim sFecha As String
    Dim oConec As DConecta
    
    sFecha = Format(pdFecha, "YYYYMMdd")
'Disponible
    sSQL = " select 'Caja' as Cta ,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as Caja"
    sSQL = sSQL & " from ctasaldo where (cctacontcod  like '11110[12]%' or cctacontcod  like '111701%') and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Caja' as Cta ,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as Caja"
    sSQL = sSQL & " from ctasaldo where (cctacontcod  like '11210[12]%' or cctacontcod  like '112701%') and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
'Dep. en cta Cte
    sSQL = sSQL & " select 'Dep. en Cta Cte.' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [Dep. en Cta Cte.]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '111301__01%' and len(cctacontcod)=10 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Dep. en Cta Cte.' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [Dep. en Cta Cte.]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '112301__01%' and len(cctacontcod)=10 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
'Depositos Ahorro
    sSQL = sSQL & " select 'Depósitos Ahorro' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [Depósitos Ahorro]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '111301__02%' and len(cctacontcod)=10 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Depósitos Ahorro' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [Depósitos Ahorro]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '112301__02%' and len(cctacontcod)=10 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
'Dep. Plazo Fijo
    sSQL = sSQL & " select 'Dep. a plazo' as Cuentas,'Soles' as cmoneda,isnull(sum(nctasaldoimporte),0) as [Dep. a plazo]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '111301__03%' and len(cctacontcod)=10 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Dep. a plazo' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [Dep. a plazo]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '112301__03%' and len(cctacontcod)=10 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
'Dep. CMAC Plazo Fijo y Aho

    sSQL = sSQL & " select 'Dep. en CMAC (Aho y Plazo)' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [Dep. en CMAC (Aho y Plazo)]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '111303__%' and len(cctacontcod)=10 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Dep. en CMAC (Aho y Plazo)' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [Dep. en CMAC (Aho y Plazo)]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '112303__%' and len(cctacontcod)=10 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
Set oConec = New DConecta
oConec.AbreConexion
Set Recup_Caja = oConec.CargaRecordSet(sSQL)
oConec.CierraConexion
Set oConec = Nothing
End Function

Public Function Recup_Adeudados(ByVal pdFecha As Date, ByVal pnTipoCambio As Double) As Recordset
    Dim sSQL As String
    Dim oConec As DConecta
    Dim sFecha As String
    '-Adeudados

    sFecha = Format(pdFecha - 29, "YYYYMMdd")
    
    sSQL = "select 'Prest. Bancos Wiesse Sud' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '241302010105__%' and len(cctacontcod)=12 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prest. Bancos Wiesse Sud' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '242302010105__%' and len(cctacontcod)=12 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prest. Bancos Credito' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '241302010103__%' and len(cctacontcod)=12 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prest. Bancos Credito' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '242302010103__%' and len(cctacontcod)=12 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prest. Bancos AgroBanco' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '241302010131__%' and len(cctacontcod)=12 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prest. Bancos AgroBanco' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '242302010131__%' and len(cctacontcod)=12 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prest. Bancos COFIDE' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where (cctacontcod like '24120[23]0101%'or cctacontcod like '26120201020[12]%') and  len(cctacontcod)=10 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prest. Bancos COFIDE' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where (cctacontcod like '241302010131__%' or cctacontcod like '26220201020[12]%') and len(cctacontcod)=12 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prestamos de FOCMAC' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '241302010130__%' and len(cctacontcod)=12 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prestamos de FOCMAC' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '242302010130__%' and len(cctacontcod)=12 and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prestamos de BID' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '26150901%' and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prestamos de BID' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [nValor]"
    sSQL = sSQL & "from ctasaldo where cctacontcod like '26250901%' and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prestamos de FONCODES' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '2416010109%' and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prestamos de FONCODES' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '2426010109%' and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prestamos de Otra CMAC' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '241601010911%' and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Prestamos de Otra CMAC' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '242601010911%' and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Fondos en Fideicomiso' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '241601010911%' and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Fondos en Fideicomiso' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '242601010911%' and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set Recup_Adeudados = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function Recup_Gastos(ByVal dFecha As Date, ByVal pnTipoCambio As Double) As Recordset
    Dim sSQL As String
    Dim oConec As DConecta
    Dim sFecha As String
    sFecha = Format(dFecha, "YYYYMMdd")
    'Gastos
    sSQL = "Select 'Gastos de Personal' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '451[12]%' and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Gastos de Personal' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '452[12]%' and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Gastos Administrativos' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '4513%' and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Gastos Administrativos' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [nValor]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '4513%' and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set Recup_Gastos = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
End Function

Public Function Recup_Patrimonio(ByVal dFecha As Date, ByVal pnTipoCambio As Double) As Recordset
    Dim sSQL As String
    Dim oConec As DConecta
    Dim sFecha As String
    sFecha = Format(dFecha - 29, "YYYYMMdd")
'Patrimonio
    sSQL = "select 'Patrimonio' as Cuentas,'Soles' as cmoneda, isnull(sum(nctasaldoimporte),0) as [Patrimonio]"
    sSQL = sSQL & " from ctasaldo where cctacontcod like '3%'  and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Patrimonio' as Cuentas,'Dolares' as cmoneda, isnull(sum(nctasaldoimporteME*" & pnTipoCambio & "),0) as [Patrimonio]"
    sSQL = sSQL & " from ctasaldo where cctacontcod = '3'  and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set Recup_Patrimonio = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function Recup_Inversiones(ByVal dFecha As Date, ByVal pnTipoCambio As Double) As Recordset

    Dim sSQL As String
    Dim oConec As DConecta
    Dim sFecha As String
    sFecha = Format(dFecha - 29, "YYYYMMdd")
    sSQL = " select 'Inversiones' as Cuentas,'Soles' as cmoneda,isnull(sum(nctasaldoimporte),0) as nValor from ctasaldo"
    sSQL = sSQL & " where cctacontcod like '17%' and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select 'Inversiones' as Cuentas,'Dolares' as cmoneda,isnull(sum(nctasaldoimporte),0) as nValor from ctasaldo"
    sSQL = sSQL & " where cctacontcod = '17' and convert(varchar(8),dctasaldofecha,112)='" & sFecha & "'"

    Set oConec = New DConecta
    oConec.AbreConexion
    Set Recup_Inversiones = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing

End Function

Public Function Recup_Personal(ByVal dFecha As Date, ByVal pnTipoCambio As Double) As Recordset
    Dim sSQL As String
    Dim oConec As DConecta
    Dim sFecha As String
    sFecha = Format(dFecha, "YYYYMMdd")
    
    sSQL = "select count(*) as Personal from rrhh where nrhestado ='201' and cuser not in ('XXXX','') AND convert(varchar(8),dingreso,112)<='" & sFecha & "'"
 
    Set oConec = New DConecta
    oConec.AbreConexion
    Set Recup_Personal = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

