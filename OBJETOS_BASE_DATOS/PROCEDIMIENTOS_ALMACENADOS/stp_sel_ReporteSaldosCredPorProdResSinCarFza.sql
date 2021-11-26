create procedure stp_sel_ReporteSaldosCredPorProdResSinCarFza
(
	@nTipCam money
)
As
Begin
	select substring(convert(char(3),nconsvalor),1,1) TipProd,min(Producto)Producto,sum(xy.SaldoTot)SaldoTot 
	from	(
	select	convert(char(3),c.nconsvalor) nconsvalor,max(c.cConsDescripcion) Producto,
		isnull(SUM(case substring(cc.cCtaCod,9,1) when '1' then cc.nSaldoCap else cc.nSaldoCap * @nTipCam end),0.00) SaldoTot
	from constante c
	left join DBConsolidada..CreditoConsol cc on substring(cc.cctacod,6,3)=c.nconsvalor
	where c.nconscod=1001 and c.nconsvalor not in (230,232,233,234,121,221)
	group by c.nconsvalor
	) xy 
	group by substring(convert(char(3),nconsvalor),1,1)
	order by substring(convert(char(3),nconsvalor),1,1)
End
