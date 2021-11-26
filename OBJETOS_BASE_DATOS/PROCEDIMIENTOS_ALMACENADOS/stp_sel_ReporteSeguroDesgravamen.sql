ALTER procedure stp_sel_ReporteSeguroDesgravamen
(
	@TipCam float,
	@FecAct datetime
)
as
Begin
	select C.cCtaCod,substring(C.cCtaCod,4,2) as cCodOfi,P.cPersCod,P.nPersPersoneria,CDTS.nNroCred,P.cPersNombre,
		(select top 1 cPersIdNro from DBCmacMaynas..PersId where cPersCod=P.cPersCod order by cPersIdTpo) as cPersIdNro,
		C.nMontoDesemb,C.nCuotasApr,C.dFecVig,C.nSaldoCap,
		(case when C.dFecVig>='20060817' then isnull(TIN.nInteres,0) else 0 end) as nInteres,
		(case when C.dFecVig>='20060817' then C.nSaldoCap+isnull(TIN.nInteres,0) else C.nSaldoCap end) as Total,
		(case when substring(C.cCtaCod,9,1)='1' 
			then C.nSaldoCap +  case when C.dFecVig>='20060817' then isnull(TIN.nInteres,0) else 0 end
			else (C.nSaldoCap + case when C.dFecVig>='20060817' then isnull(TIN.nInteres,0) else 0 end)*3.00 
		end) as CAPIMN,
		P.dPersNacCreac,(datediff(day,P.dPersNacCreac,@FecAct)/365.25) as cEdad,P.cPersDireccDomicilio,
		(case when substring(C.cCtaCod,9,1)='1' then 'MN' else 'ME' end) as cMoneda,
		(case substring(C.cCtaCod,6,2)
			when '10' then case when C.nDiasAtraso <=15 then 'N' else 'X' end --Comercial
			when '20' then case when C.nDiasAtraso <=30 then 'N' else 'X' end --Mes
			when '30' then case when C.nDiasAtraso <=30 then 'N' else 'X' end --Consumo
			when '32' then case when C.nDiasAtraso <=30 then 'N' else 'X' end --Consumo trabajadores
			when '40' then case when C.nDiasAtraso <=30 then 'N' else 'X' end --Hipotecario
			else 'X'
		end) as cContab,
		(case isnull(CC.nExoSeguroDes,2) 
			when 0 then 'S'
			when 1 then 'N'
			else ''
		end) as cSgrDsg,
		isnull(COD.cPersCod,'') as Codeudor,isnull(COD.cPersNombre,'') as cNomCli_Cod,isnull(COD.cPersIdNro,'') as cNuDoCi_Cod,COD.dPersNacCreac as dNacimi_Cod,
		isnull(CON.cPersCod,'') as Conyugue,isnull(CON.cPersNombre,'') as cNomCli_Con,isnull(CON.cPersIdNro,'') as cNuDoCi_Con,CON.dPersNacCreac as dNacimi_Con,
		isnull(REP.cPersCod,'') as Representante,isnull(REP.cPersNombre,'') as cNomCli_Rep,isnull(REP.cPersIdNro,'') as cNuDoCi_Rep,REP.dPersNacCreac as dNacimi_Rep
	from DBConsolidada..CreditoConsol C 
		inner join DBConsolidada..ProductoPersonaConsol PP on C.cCtaCod=PP.cCtaCod and PP.nPrdPersRelac=20
		inner join DBConsolidada..Persona P on PP.cPersCod=P.cPersCod
		inner join 
			(select PPN.cPersCod,count(PPN.cPersCod) as nNroCred 
			from DBConsolidada..CreditoConsol PN 
				inner join DBConsolidada..ProductoPersonaConsol PPN on PN.cCtaCod=PPN.cCtaCod
			where PN.nPrdEstado in (2020,2021,2022,2030,2031,2032) and substring(PN.cCtaCod,6,3) not in ('121','221','305') and PPN.nPrdPersRelac=20 
			group by PPN.cPersCod
			) CDTS on P.cPersCod=CDTS.cPersCod
		left join 
			(select PP2.cCtaCod,P2.cPersCod,P2.cPersNombre,P2.dPersNacCreac,
				(select top 1 cPersIdNro from DBCmacMaynas..PersId where cPersCod=P2.cPersCod) as cPersIdNro
			from DBConsolidada..ProductoPersonaConsol PP2 
				inner join DBConsolidada..Persona P2 on PP2.cPersCod=P2.cPersCod and PP2.nPrdPersRelac=22 --Codeudor
			) as COD on C.cCtaCod=COD.cCtaCod
		left join
			(select PP2.cCtaCod,P2.cPersCod,P2.cPersNombre,P2.dPersNacCreac,
				(select top 1 cPersIdNro from DBCmacMaynas..PersId where cPersCod=P2.cPersCod) as cPersIdNro
			from DBConsolidada..ProductoPersonaConsol PP2 
				inner join DBConsolidada..Persona P2 on PP2.cPersCod=P2.cPersCod and PP2.nPrdPersRelac=21 --Conyugue
			) as CON on C.cCtaCod=CON.cCtaCod
		left join
			(select PP2.cCtaCod,P2.cPersCod,P2.cPersNombre,P2.dPersNacCreac,
				(select top 1 cPersIdNro from DBCmacMaynas..PersId where cPersCod=P2.cPersCod) as cPersIdNro
			--from DBConsolidada..ProductoPersonaConsol PP2 inner join DBConsolidada..Persona P2 on PP2.cPersCod=P2.cPersCod and PP2.nPrdPersRelac=23 --Representante
			from ProductoPersona PP2 inner join DBConsolidada..Persona P2 on PP2.cPersCod=P2.cPersCod and PP2.cSgrDsg='S'
			) as REP on C.cCtaCod=REP.cCtaCod
		left join 
			(select PPC.cCtaCod,PPC.nInteres 
			from DBConsolidada..PlanDesPagConsol PPC
				inner join 
				(select cCtaCod,min(cNroCuo) as cNroCuo from DBConsolidada..PlanDesPagConsol 
				where nEstado=0 and nTipo=1 group by cCtaCod) as T on PPC.cCtaCod=T.cCtaCod and PPC.cNroCuo=T.cNroCuo and PPC.nTipo=1
			) as TIN on C.cCtaCod=TIN.cCtaCod 
		left join DBCmacMaynas..ColocacCred CC on C.cCtaCod=CC.cCtaCod
	where C.nPrdEstado in (2020,2021,2022,2030,2031,2032) --Vigentes
		and substring(C.cCtaCod,6,3) not in ('121','221','305') --Excluir cartas fianzas y pignoraticio
		and C.nDiasAtraso <=(case substring(C.cCtaCod,6,2) 
								when '10' then 15 
								when '20' then 30
								when '30' then 30
								WHEN '32' THEN 30
								when '40' then 30 end)
	order by substring(C.cCtaCod,4,2)
End