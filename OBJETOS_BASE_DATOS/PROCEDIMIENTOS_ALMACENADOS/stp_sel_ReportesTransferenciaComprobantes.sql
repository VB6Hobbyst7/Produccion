create procedure stp_sel_ReportesTransferenciaComprobantes
	@cPeri varchar(6),
	@cCtaIgv varchar(10),
	@nTper int,
	@nEsCon int,
	@nMovEl int,
	@nMovEx int,
	@nMovEx2 int,
	@nMovMod int,
	@cOpeCod varchar(15),
	@cTpoDoc varchar(20)
as
begin	
	declare @uit int
    set @uit=(select nParValor*0.1 from parametro where nParProd=2001 and nParCod=1)
	if @cTpoDoc = ''  
	begin
		SELECT M.cMovNro,M.nMovNro,M.cOpeCod, MD.nDocTpo,ISNULL(P.cPersNombre,'') cPersNombre, ISNULL(pid.cPersIDnro,'') as cProvRuc,MD.dDocFecha,  MD.cDocNro, ISNULL(SUM(MOTI.nMovOtroItemImporte),0) nIGV, ISNULL(SUM(MOTO.nMovOtroItemImporte),0)  nOtros, sum(MC.nMovImporte) as nMovImporte, SUM(me.nMovMEImporte) as nMovMEImporte, ISNULL(mtc.nMovTpoCambio,1) nTC  
		FROM MOV M JOIN MOVCTA MC ON M.nMovNro = MC.nMovNro 
					LEFT JOIN MovME me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem            
		JOIN MOVDOC MD ON M.nMovNro = MD.nMovNro       
		LEFT JOIN (SELECT nMovNro, nMovItem, SUM(nMovOtroImporte) as nMovOtroItemImporte                  
					FROM MOVOTROSITEM WHERE cMovOtroVariable = @cCtaIgv                  
					GROUP BY nMovNro, nMovItem                 ) AS MOTI ON MC.nMovNro = MOTI.nMovNro and MC.nMovItem = MOTI.nMovItem       
		LEFT JOIN (SELECT nMovNro, nMovItem, SUM(nMovOtroImporte) as nMovOtroItemImporte                  
					FROM MOVOTROSITEM WHERE not cMovOtroVariable = @cCtaIgv                  
					GROUP BY nMovNro, nMovItem) AS MOTO ON MC.nMovNro = MOTO.nMovNro and MC.nMovItem = MOTO.nMovItem            
		JOIN MovGasto MO ON mo.nMovNro = m.nMovNro            
		JOIN OpeDoc OD ON od.nDocTpo = md.nDocTpo       
		LEFT JOIN MovTpoCambio mtc ON mtc.nMovNro = m.nMovNro       
		LEFT JOIN Persona P ON MO.cPersCod = P.cPersCod       
		LEFT JOIN PersID pid ON pid.cPersCod = P.cPersCod and pid.cPersIDTpo = @nTper 
		WHERE M.nMovEstado = @nEsCon and M.nMovFlag not IN (@nMovEl,@nMovEx,@nMovEx2,@nMovMod)   and Substring(M.cMovNro,1,6) LIKE @cPeri + '%' 
		and OD.cOpeCod = @cOpeCod and MC.nMovImporte > 0 
		GROUP BY M.cMovNro,M.nMovNro,M.cOpeCod,MD.nDocTpo,P.cPersNombre, pid.cPersIDnro,MD.dDocFecha, MD.cDocNro, ISNULL(mtc.nMovTpoCambio,1) 
		HAVING SUM(MC.nMovImporte) >= @uit 
		ORDER BY P.cPersNombre, md.cDocNro 
end
else
begin
		SELECT M.cMovNro,M.nMovNro,M.cOpeCod, MD.nDocTpo,ISNULL(P.cPersNombre,'') cPersNombre, ISNULL(pid.cPersIDnro,'') as cProvRuc,MD.dDocFecha,  MD.cDocNro, ISNULL(SUM(MOTI.nMovOtroItemImporte),0) nIGV, ISNULL(SUM(MOTO.nMovOtroItemImporte),0)  nOtros, sum(MC.nMovImporte) as nMovImporte, SUM(me.nMovMEImporte) as nMovMEImporte, ISNULL(mtc.nMovTpoCambio,1) nTC  
		FROM MOV M JOIN MOVCTA MC ON M.nMovNro = MC.nMovNro 
					LEFT JOIN MovME me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem            
		JOIN MOVDOC MD ON M.nMovNro = MD.nMovNro       
		LEFT JOIN (SELECT nMovNro, nMovItem, SUM(nMovOtroImporte) as nMovOtroItemImporte                  
					FROM MOVOTROSITEM WHERE cMovOtroVariable = @cCtaIgv                  
					GROUP BY nMovNro, nMovItem                 ) AS MOTI ON MC.nMovNro = MOTI.nMovNro and MC.nMovItem = MOTI.nMovItem       
		LEFT JOIN (SELECT nMovNro, nMovItem, SUM(nMovOtroImporte) as nMovOtroItemImporte                  
					FROM MOVOTROSITEM WHERE not cMovOtroVariable = @cCtaIgv                  
					GROUP BY nMovNro, nMovItem) AS MOTO ON MC.nMovNro = MOTO.nMovNro and MC.nMovItem = MOTO.nMovItem            
		JOIN MovGasto MO ON mo.nMovNro = m.nMovNro            
		JOIN OpeDoc OD ON od.nDocTpo = md.nDocTpo       
		LEFT JOIN MovTpoCambio mtc ON mtc.nMovNro = m.nMovNro       
		LEFT JOIN Persona P ON MO.cPersCod = P.cPersCod       
		LEFT JOIN PersID pid ON pid.cPersCod = P.cPersCod and pid.cPersIDTpo = @nTper 
		WHERE M.nMovEstado = @nEsCon and M.nMovFlag not IN (@nMovEl,@nMovEx,@nMovEx2,@nMovMod)   and Substring(M.cMovNro,1,6) LIKE @cPeri + '%' 
		and OD.cOpeCod = @cOpeCod and MC.nMovImporte > 0 
		and md.nDocTpo IN (select valor from dbo.fnc_getTblValoresNumerico(@cTpoDoc))
		GROUP BY M.cMovNro,M.nMovNro,M.cOpeCod,MD.nDocTpo,P.cPersNombre, pid.cPersIDnro,MD.dDocFecha, MD.cDocNro, ISNULL(mtc.nMovTpoCambio,1) 
		HAVING SUM(MC.nMovImporte) >= @uit 
		ORDER BY P.cPersNombre, md.cDocNro 
	end
end
