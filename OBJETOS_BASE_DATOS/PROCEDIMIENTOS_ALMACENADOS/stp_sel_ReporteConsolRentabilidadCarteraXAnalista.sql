create procedure stp_sel_ReporteConsolRentabilidadCarteraXAnalista
(
	@dfechaIni datetime,
	@dfechaFin datetime,
	@ndiaMora int,
	@sAnalistas varchar(max),
	@sAgencias varchar(max)
)
as
begin
	select	CodAgencia,max(Agencia)Agencia,cUser,max(cNomUser)cNomUser,cMoney,UPPER(max(cMoneda))cMoneda,
			count(cCtaCod) Casos,sum(nmontocol) nmontocol,sum(nsaldo) nsaldo,sum(nmontocuota)nmontocuota,
			sum(capital)capital,sum(interes)interes,sum(Mora)Mora,sum(gastos)gastos
		from (
		select substring(p.cctacod,4,2) CodAgencia,upper(cAgeDescripcion) Agencia,upper(RH.cUser)cUser,
		upper(per.cPersNombre)cNomUser,substring(P.cCtaCod,9,1) cMoney,P.cCtaCod,
		case substring(P.cCtaCod,9,1) when 1 then 'Soles' else 'Dolares' end cMoneda,
		C.nMontoCol,P.nSaldo,
		nMontoCuota=isnull((Select Sum(nMontoPagado) From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl and nCuota=CC.nCuota and (nPrdConceptoCod in (1000,1100,1101,1102,1103) or nPrdConceptoCod like '12%')),0) ,
		Capital=isnull((Select Sum(nMontoPagado) From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl and nCuota=CC.nCuota and nPrdConceptoCod in (1000)),0) ,
		Interes=isnull((Select Sum(nMontoPagado) From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl and nCuota=CC.nCuota and nPrdConceptoCod in (1100,1102,1103)),0),
		Mora=isnull((Select Sum(nMontoPagado) From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl and nCuota=CC.nCuota and nPrdConceptoCod in (1101)),0),
		Gastos=isnull((Select Sum(nMontoPagado) From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl and nCuota=CC.nCuota and nPrdConceptoCod like '12%'),0)
		from producto P inner join Colocaciones C on P.cCtaCod=C.cCtaCod and len(C.cLineaCred)=13
			inner join agencias ag on substring(p.cctacod,4,2)=ag.cagecod
			inner join Productopersona PP on (P.cCtaCod=PP.cCtaCod) and PP.nPrdPersRelac=20
			inner join colocCalendario CC on (CC.cCtaCod=C.cCtaCod) and CC.nColocCalendApl=1 and CC.nColocCalendEstado=1		
			inner join ProductoPersona PP2 on (P.cCtaCod=PP2.cCtaCod) and PP2.nPrdPersRelac=28
			inner join RRHH RH on (PP2.cPersCod=RH.cPersCod)
			inner join Persona per on rh.cperscod=per.cperscod
			inner join constante Cons1 on (Cons1.nConsValor=substring(P.cCtaCod,6,3)) and nConsCod=1001
			inner join (select Valor from fnc_getTblValoresTexto (@sAnalistas)) AnaL on (AnaL.Valor=RH.cPersCod)
		where  datediff(day,CC.dvenc,CC.dPago)<=@ndiaMora and CC.dPago between @dfechaIni and @dfechaFin and 
			  CC.nNroCalen=(select max(nNroCalen) from colocCalendario where cCtaCod=CC.cCtaCod)
			and Substring(P.cCtaCod,4,2) in (select Valor from fnc_getTblValoresTexto (@sAgencias)) 
	) xy
	group by xy.CodAgencia,xy.cUser,xy.cMoney
	order by xy.CodAgencia,xy.cUser,xy.cMoney
end
