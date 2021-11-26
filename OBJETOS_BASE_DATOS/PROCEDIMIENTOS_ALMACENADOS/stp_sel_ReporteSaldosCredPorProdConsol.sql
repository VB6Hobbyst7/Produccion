create procedure stp_sel_ReporteSaldosCredPorProdConsol
(
	@nTipCam money
)
As
Begin
	select substring(convert(char(3),nconsvalor),1,1) TipProd,xy.* 
	from	(
	select	convert(char(3),c.nconsvalor) nconsvalor,max(c.cConsDescripcion) Producto,
		isnull(count(case when substring(cc.cctacod,9,1) ='1' then cc.cctacod end),0.00) NumCredMN,
		isnull(count(case when substring(cc.cctacod,9,1) ='2' then cc.cctacod end),0.00) NumCredME,
		isnull(SUM(case when substring(cc.cCtaCod,9,1) = '1' then cc.nSaldoCap else 0.00 end),0.00) SaldoMN,
		isnull(SUM(case when substring(cc.cCtaCod,9,1) = '2' then cc.nSaldoCap else 0.00 end),0.00) SaldoME,
		isnull(SUM(case substring(cc.cCtaCod,9,1) when '1' then cc.nSaldoCap else cc.nSaldoCap * @nTipCam end),0.00) SaldoTot
	from constante c
	left join DBConsolidada..CreditoConsol cc on substring(cc.cctacod,6,3)=c.nconsvalor
	where c.nconscod=1001 and c.nconsvalor not in (230,232,233,234,121,221)
	group by c.nconsvalor
	union
	select	convert(char(3),c.nconsvalor) nconsvalor,max(c.cConsDescripcion) Producto,
		isnull(count(case when substring(cc.cctacod,9,1) ='1' then cc.cctacod end),0.00) NumCredMN,
		isnull(count(case when substring(cc.cctacod,9,1) ='2' then cc.cctacod end),0.00) NumCredME,
		isnull(SUM(case when substring(cc.cCtaCod,9,1) = '1' then cc.nSaldoCap else 0.00 end),0.00) SaldoMN,
		isnull(SUM(case when substring(cc.cCtaCod,9,1) = '2' then cc.nSaldoCap else 0.00 end),0.00) SaldoME,
		isnull(SUM(case substring(cc.cCtaCod,9,1) when '1' then cc.nSaldoCap else cc.nSaldoCap * @nTipCam end),0.00) SaldoTot
	from constante c
	left join DBConsolidada..CartaFianzaSaldoConsol cc on substring(cc.cctacod,6,3)=c.nconsvalor
	where c.nconscod=1001 and c.nconsvalor in (121,221)
	group by c.nconsvalor
	) xy order by xy.nconsvalor
End
