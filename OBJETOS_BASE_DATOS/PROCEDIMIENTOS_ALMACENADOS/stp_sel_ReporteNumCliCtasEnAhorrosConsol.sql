create procedure stp_sel_ReporteNumCliCtasEnAhorrosConsol As
Begin
	select	xy.CodMone, case when xy.CodMone='1' then 'SOLES  ' else 'DOLARES' end NomMoneda,
			sum(case when TipoAho ='1' then xy.NumCliente else 0 end) NumCliAho,
			sum(case when TipoAho ='1' then xy.NumCuentas else 0 end) NumCtaAho,
			sum(case when TipoAho ='2' then xy.NumCliente else 0 end) NumCliPla,
			sum(case when TipoAho ='2' then xy.NumCuentas else 0 end) NumCtaPla,
			sum(case when TipoAho ='3' then xy.NumCliente else 0 end) NumCliCts,
			sum(case when TipoAho ='3' then xy.NumCuentas else 0 end) NumCtaCts
	from (
	select	'1' TipoAho, substring(aho.cctacod,9,1) CodMone,
			count(distinct pp.cperscod) NumCliente,
			count(pp.cctacod) NumCuentas
		from DBConsolidada..AhorroCConsol aho
		inner join productopersona pp on aho.cctacod=pp.cctacod and pp.nPrdPersRelac=10
		where substring(aho.cctacod,9,1) in (1,2)
	group by substring(aho.cctacod,9,1)
	union
	select	'2' TipoAho, substring(pla.cctacod,9,1) CodMone,
			count(distinct pp.cperscod) NumCliente,
			count(pp.cctacod) NumCuentas
		from DBConsolidada..PlazoFijoConsol pla
		inner join productopersona pp on pla.cctacod=pp.cctacod and pp.nPrdPersRelac=10
	group by substring(pla.cctacod,9,1)
	union
	select	'3' TipoAho, substring(cts.cctacod,9,1) CodMone,
			count(distinct pp.cperscod) NumCliente,
			count(pp.cctacod) NumCuentas
		from DBConsolidada..CTSConsol cts
		inner join productopersona pp on cts.cctacod=pp.cctacod and pp.nPrdPersRelac=10
	group by substring(cts.cctacod,9,1)
	) xy
	group by xy.CodMone
End
