create procedure stp_sel_ReporteGetITFQuincena
@sFechaIni varchar(8),
@sFechaFin varchar(8),
@nTipoReporte int

as
if @nTipoReporte=1
begin
		Select Oficina,  
			IsNull(Sum(Case When aaa.cOpecod In ('990101','990102','990109','990301','990302') Then nMontoDetSol End),0) MontoSolAho,  
			IsNull(Sum(Case When aaa.cOpecod In ('990101','990102','990109','990301','990302') Then nMontoDetDol End),0) As MonDolAho,  
			IsNull(Sum(Case When aaa.cOpecod In ('990103','990105','990106','990107','990108','990303') Then nMontoDetSol End),0) As MonSolCre,  
			IsNull(Sum(Case When aaa.cOpecod In ('990103','990105','990106','990107','990108','990303') Then nMontoDetDol End),0) As MonDolCre,  
			IsNull(Sum(Case When aaa.cOpecod In ('990104','990304') Then nMontoDetSol End),0) As MonSolPre,   
			IsNull(Sum(Case When aaa.cOpecod In ('1') Then nMontoDetSol End),0) As MonSolCaj,   
			IsNull(Sum(Case When aaa.cOpecod In ('2') Then nMontoDetDol End),0) As MonDolCaj   
		From ( 
				SELECT  MC.CCTACOD, 
				(SELECT max(cperscod) FROM ProductoPersona pp where pp.cctacod = MC.cctacod And nPrdPersRelac = 20) CodPers ,     
					Substring(MC.CCTACOD,4,2) oficina,MC.COPECOD,O.cOpeDesc,     
					SUM(Case Substring(MC.CCTACOD,9,1) when '2' then  round(MD.nMonto,2) else 0.00 end) as nMontoDetDol ,     
					SUM(Case Substring(MC.CCTACOD,9,1) when '1' then  Round(MD.nMonto,2) else 0.00 end) as nMontoDetSol  
				FROM MOV M		JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO     
								JOIN MOVCOLDET MD ON MD.NMOVNRO = MC.NMOVNRO	AND MD.COPECOD = MC.COPECOD 
																				AND MD.CCTACOD = MC.CCTACOD     
								JOIN OPETPO O ON O.COPECOD = MC.COPECOD     
								JOIN PRODUCTOCONCEPTO PC ON PC.nPrdConceptoCod = MD.nPrdConceptoCod  
				WHERE MC.COPECOD LIKE '99%' AND LEFT(M.CMOVNRO,8) BETWEEN @sFechaIni AND @sFechaFin AND M.NMOVFLAG = 0  
				GROUP BY M.cMovNro,MC.COPECOD,O.cOpeDesc,MC.CCTACOD  
				Union All  
				SELECT  MC.CCTACOD, 
				(SELECT max(cperscod) FROM ProductoPersona pp where pp.cctacod = MC.cctacod And nPrdPersRelac = 10) CodPers,  
					Substring(MC.CCTACOD,4,2) oficina, MC.COPECOD,O.cOpeDesc,  
					SUM(Case Substring(MC.CCTACOD,9,1) when '2' then  round(MD.nMonto,2) else 0.00 end) as nMontDetDol ,  
					SUM(Case Substring(MC.CCTACOD,9,1) when '1' then  Round(MD.nMonto,2) else 0.00 end) as nMontDetSol  
				FROM    MOV M   JOIN MOVCAP MC ON MC.NMOVNRO = M.NMOVNRO  
								JOIN MOVCAPDET MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.COPECOD = MC.COPECOD 
																			AND MD.CCTACOD = MC.CCTACOD  
								JOIN OPETPO O ON O.COPECOD = MC.COPECOD  
								JOIN PRODUCTOCONCEPTO PC ON PC.nPrdConceptoCod = MD.nConceptoCod  
				WHERE MC.COPECOD LIKE '99%' AND LEFT(M.CMOVNRO,8) BETWEEN @sFechaIni AND @sFechaFin AND M.NMOVFLAG = 0  
				GROUP BY M.cMovNro,MC.COPECOD,O.cOpeDesc,MC.CCTACOD  
				Union All  
				SELECT  '1089800000272' cCtaCod, '1089800000272' CodPers,  
					Substring(M.cMovNro,18,2) oficina,'2' as COPECOD, 'DOLARES' As cOpeDesc,  
					SUM(round(ABS(ME.NMOVMEIMPORTE),2))  as nMontDetDol, 0.00 as nMontDetSol  
				FROM    MOV M   JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO  
								LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO and ME.nMovItem = MC.nMovItem  
								JOIN OPETPO O ON O.COPECOD = M.COPECOD  WHERE 
								LEFT(M.CMOVNRO,8) BETWEEN @sFechaIni AND @sFechaFin  AND M.NMOVFLAG = 0 
																	AND M.COPECOD Not In ('701105') 
																	AND MC.cCtaContCod LIKE '21240207%'   
																	and MC.nMovImporte < 0  
				Group By M.cMovNro  
				Union All  
				SELECT '1089800000272' cCtaCod, '1089800000272' CodPers,  
					Substring(M.cMovNro,18,2) oficina,'1' as COPECOD, 'SOLES' As cOpeDesc,  0.00 As nMontDetDol,
					SUM(round(ABS(MC.NMOVIMPORTE),2))  as nMontDetSol   
				FROM    MOV M  JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO  
						LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO and ME.nMovItem = MC.nMovItem  
						JOIN OPETPO O ON O.COPECOD = M.COPECOD  
				WHERE LEFT(M.CMOVNRO,8) BETWEEN @sFechaIni AND @sFechaFin  AND M.NMOVFLAG = 0 
												AND M.COPECOD Not In ('701105') 
												AND MC.cCtaContCod LIKE '21140207%'   
												and MC.nMovImporte < 0 
				Group By M.cMovNro) As aaa  
			left join persona pe on aaa.codpers =  pe.cperscod  
			left join persid bb on aaa.codpers =  bb.cperscod and bb.cPersIDTpo = 1  
			left join persid cc on aaa.codpers =  cc.cperscod and cc.cPersIDTpo = 2  
			left join persid dd on aaa.codpers =  dd.cperscod and dd.cPersIDTpo = 4  
			left join persid ee on aaa.codpers =  ee.cperscod and ee.cPersIDTpo = 11  
			left join persid ff on aaa.codpers =  ff.cperscod and ff.cPersIDTpo = 10  
		Group by  oficina Order By oficina 
end
else
begin
		Select Oficina,  
			IsNull(Sum(Case When aaa.cOpecod In ('990101','990102','990109','990301','990302') Then nMontoDetSol End),0) MontoSolAho,  
			IsNull(Sum(Case When aaa.cOpecod In ('990101','990102','990109','990301','990302') Then nMontoDetDol End),0) As MonDolAho,  
			IsNull(Sum(Case When aaa.cOpecod In ('990103','990105','990106','990107','990108','990303') Then nMontoDetSol End),0) As MonSolCre,  
			IsNull(Sum(Case When aaa.cOpecod In ('990103','990105','990106','990107','990108','990303') Then nMontoDetDol End),0) As MonDolCre,  
			IsNull(Sum(Case When aaa.cOpecod In ('990104','990304') Then nMontoDetSol End),0) As MonSolPre,   
			IsNull(Sum(Case When aaa.cOpecod In ('1') Then nMontoDetSol End),0) As MonSolCaj,   
			IsNull(Sum(Case When aaa.cOpecod In ('2') Then nMontoDetDol End),0) As MonDolCaj   
		From ( 
				SELECT  MC.CCTACOD, 
				(SELECT max(cperscod) FROM ProductoPersona pp where pp.cctacod = MC.cctacod And nPrdPersRelac = 20) CodPers ,     
					Substring(MC.CCTACOD,4,2) oficina,MC.COPECOD,O.cOpeDesc,     
					SUM(Case Substring(MC.CCTACOD,9,1) when '2' then  round(dbo.GetTCPonderado(LEFT(M.CMOVNRO,8)) * MD.nMonto,2) else 0.00 end) as nMontoDetDol ,     
					SUM(Case Substring(MC.CCTACOD,9,1) when '1' then  Round(MD.nMonto,2) else 0.00 end) as nMontoDetSol  
				FROM MOV M		JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO     
								JOIN MOVCOLDET MD ON MD.NMOVNRO = MC.NMOVNRO	AND MD.COPECOD = MC.COPECOD 
																				AND MD.CCTACOD = MC.CCTACOD     
								JOIN OPETPO O ON O.COPECOD = MC.COPECOD     
								JOIN PRODUCTOCONCEPTO PC ON PC.nPrdConceptoCod = MD.nPrdConceptoCod  
				WHERE MC.COPECOD LIKE '99%' AND LEFT(M.CMOVNRO,8) BETWEEN @sFechaIni AND @sFechaFin AND M.NMOVFLAG = 0  
				GROUP BY M.cMovNro,MC.COPECOD,O.cOpeDesc,MC.CCTACOD  
				Union All  
				SELECT  MC.CCTACOD, 
				(SELECT max(cperscod) FROM ProductoPersona pp where pp.cctacod = MC.cctacod And nPrdPersRelac = 10) CodPers,  
					Substring(MC.CCTACOD,4,2) oficina, MC.COPECOD,O.cOpeDesc,  
					SUM(Case Substring(MC.CCTACOD,9,1) when '2' then  round(dbo.GetTCPonderado(LEFT(M.CMOVNRO,8)) * MD.nMonto,2) else 0.00 end) as nMontDetDol ,  
					SUM(Case Substring(MC.CCTACOD,9,1) when '1' then  Round(MD.nMonto,2) else 0.00 end) as nMontDetSol  
				FROM    MOV M   JOIN MOVCAP MC ON MC.NMOVNRO = M.NMOVNRO  
								JOIN MOVCAPDET MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.COPECOD = MC.COPECOD 
																			AND MD.CCTACOD = MC.CCTACOD  
								JOIN OPETPO O ON O.COPECOD = MC.COPECOD  
								JOIN PRODUCTOCONCEPTO PC ON PC.nPrdConceptoCod = MD.nConceptoCod  
				WHERE MC.COPECOD LIKE '99%' AND LEFT(M.CMOVNRO,8) BETWEEN @sFechaIni AND @sFechaFin AND M.NMOVFLAG = 0  
				GROUP BY M.cMovNro,MC.COPECOD,O.cOpeDesc,MC.CCTACOD  
				Union All  
				SELECT  '1089800000272' cCtaCod, '1089800000272' CodPers,  
					Substring(M.cMovNro,18,2) oficina,'2' as COPECOD, 'DOLARES' As cOpeDesc,  
					SUM(round(ABS(dbo.GetTCPonderado(LEFT(M.CMOVNRO,8)) * ME.NMOVMEIMPORTE),2))  as nMontDetDol, 0.00 as nMontDetSol  
				FROM    MOV M   JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO  
								LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO and ME.nMovItem = MC.nMovItem  
								JOIN OPETPO O ON O.COPECOD = M.COPECOD  WHERE 
								LEFT(M.CMOVNRO,8) BETWEEN @sFechaIni AND @sFechaFin  AND M.NMOVFLAG = 0 
																	AND M.COPECOD Not In ('701105') 
																	AND MC.cCtaContCod LIKE '21240207%'   
																	and MC.nMovImporte < 0  
				Group By M.cMovNro  
				Union All  
				SELECT '1089800000272' cCtaCod, '1089800000272' CodPers,  
					Substring(M.cMovNro,18,2) oficina,'1' as COPECOD, 'SOLES' As cOpeDesc,  0.00 As nMontDetDol,
					SUM(round(ABS(MC.NMOVIMPORTE),2))  as nMontDetSol   
				FROM    MOV M  JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO  
						LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO and ME.nMovItem = MC.nMovItem  
						JOIN OPETPO O ON O.COPECOD = M.COPECOD  
				WHERE LEFT(M.CMOVNRO,8) BETWEEN @sFechaIni AND @sFechaFin  AND M.NMOVFLAG = 0 
												AND M.COPECOD Not In ('701105') 
												AND MC.cCtaContCod LIKE '21140207%'   
												and MC.nMovImporte < 0 
				Group By M.cMovNro) As aaa  
			left join persona pe on aaa.codpers =  pe.cperscod  
			left join persid bb on aaa.codpers =  bb.cperscod and bb.cPersIDTpo = 1  
			left join persid cc on aaa.codpers =  cc.cperscod and cc.cPersIDTpo = 2  
			left join persid dd on aaa.codpers =  dd.cperscod and dd.cPersIDTpo = 4  
			left join persid ee on aaa.codpers =  ee.cperscod and ee.cPersIDTpo = 11  
			left join persid ff on aaa.codpers =  ff.cperscod and ff.cPersIDTpo = 10  
		Group by  oficina Order By oficina 
end
