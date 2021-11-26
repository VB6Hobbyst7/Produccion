create procedure stp_sel_ReporteNumCliCtasEnAhorros As
Begin
	select	xy.codage,upper(max(age.cAgeDescripcion)) Agencia, xy.Moneda, case when xy.Moneda='1' then 'SOLES  ' else 'DOLARES' end NomMoneda,
			isnull(sum(case when TipoAho ='1' then xy.NumCliente else 0 end),0) NumCliAho,
			isnull(sum(case when TipoAho ='1' then xy.NumCuentas else 0 end),0) NumCtaAho,
			isnull(sum(case when TipoAho ='2' then xy.NumCliente else 0 end),0) NumCliPla,
			isnull(sum(case when TipoAho ='2' then xy.NumCuentas else 0 end),0) NumCtaPla,
			isnull(sum(case when TipoAho ='3' then xy.NumCliente else 0 end),0) NumCliCts,
			isnull(sum(case when TipoAho ='3' then xy.NumCuentas else 0 end),0) NumCtaCts
	from (
	select	'1' TipoAho, substring(aho.cctacod,4,2) CodAge,substring(aho.cctacod,9,1) Moneda,
			count(distinct pp.cperscod) NumCliente,
			count(pp.cctacod) NumCuentas
		from DBConsolidada..AhorroCConsol aho
		inner join productopersona pp on aho.cctacod=pp.cctacod and pp.nPrdPersRelac=10
		where substring(aho.cctacod,9,1) in ('1','2')
	group by substring(aho.cctacod,4,2),substring(aho.cctacod,9,1)
	union
	select	'2' TipoAho, substring(pla.cctacod,4,2) CodAge,substring(pla.cctacod,9,1) Moneda,
			count(distinct pp.cperscod) NumCliente,
			count(pp.cctacod) NumCuentas
		from DBConsolidada..PlazoFijoConsol pla
		inner join productopersona pp on pla.cctacod=pp.cctacod and pp.nPrdPersRelac=10
	group by substring(pla.cctacod,4,2),substring(pla.cctacod,9,1)
	union
	select	'3' TipoAho, substring(cts.cctacod,4,2) CodAge,substring(cts.cctacod,9,1) Moneda,
			count(distinct pp.cperscod) NumCliente,
			count(pp.cctacod) NumCuentas
		from DBConsolidada..CTSConsol cts
		inner join productopersona pp on cts.cctacod=pp.cctacod and pp.nPrdPersRelac=10
	group by substring(cts.cctacod,4,2),substring(cts.cctacod,9,1)
	) xy
	Inner Join Agencias age on xy.codage=age.cAgeCod
	group by xy.codage,xy.Moneda
	order by xy.codage,xy.Moneda
End
