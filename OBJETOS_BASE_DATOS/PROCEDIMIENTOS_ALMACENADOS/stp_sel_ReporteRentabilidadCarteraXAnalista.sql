ALTER procedure stp_sel_ReporteRentabilidadCarteraXAnalista
(
	@dfechaIni datetime,
	@dfechaFin datetime,
	@ndiaMora int,
	@sAnalistas varchar(max),
	@sAgencias varchar(max)
)
as
begin
	select C.cLineaCred,CL.cDescripcion,
	Cons1.cConsDescripcion,
	case substring(P.cCtaCod,9,1) when 1 then 'Soles' else 'Dolares' end cMoneda,
	case substring(C.cLineaCred,6,1) when 1 then 'Largo Plazo' else 'Corto Plazo' end cPlazo,
	P.cCtaCod,PE.cPersNombre,C.nMontoCol,P.nSaldo,
	nMontoCuota=isnull((Select Sum(nMontoPagado) From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl and nCuota=CC.nCuota and (nPrdConceptoCod in (1000,1100,1101,1102,1103) or nPrdConceptoCod like '12%')),0) ,
	Capital=isnull((Select Sum(nMontoPagado) From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl and nCuota=CC.nCuota and nPrdConceptoCod in (1000)),0) ,
	Interes=isnull((Select Sum(nMontoPagado) From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl and nCuota=CC.nCuota and nPrdConceptoCod in (1100,1102,1103)),0),
	Mora=isnull((Select Sum(nMontoPagado) From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl and nCuota=CC.nCuota and nPrdConceptoCod in (1101)),0),
	Gastos=isnull((Select Sum(nMontoPagado) From ColocCalendDet NOLOCK Where cCtaCod = CC.cCtaCod And nNroCalen = CC.nNroCalen And nColocCalendApl = CC.nColocCalendApl and nCuota=CC.nCuota and nPrdConceptoCod like '12%'),0),
	NroCuotaApr=(Select Count(*) From ColocCalendario CD Where CD.cCtaCod=P.cCtaCod and CD.nColocCalendApl=1 and CD.nNroCalen=CC.nNroCalen),
	CC.nCuota,
	CC.dvenc,CC.dPago,RH.cUser
	from producto P inner join Colocaciones C on P.cCtaCod=C.cCtaCod and len(C.cLineaCred)=13
		inner join ColocLineaCredito CL on (C.cLineaCred=CL.cLineaCred)-- and CL.bEstado=1
		inner join Productopersona PP on (P.cCtaCod=PP.cCtaCod) and PP.nPrdPersRelac=20
		inner join Persona PE on (PP.cPersCod=PE.cPersCod)
		inner join colocCalendario CC on (CC.cCtaCod=C.cCtaCod) and CC.nColocCalendApl=1 and CC.nColocCalendEstado=1		
		inner join ProductoPersona PP2 on (P.cCtaCod=PP2.cCtaCod) and PP2.nPrdPersRelac=28
		inner join RRHH RH on (PP2.cPersCod=RH.cPersCod)
		inner join constante Cons1 on (Cons1.nConsValor=substring(P.cCtaCod,6,3)) and nConsCod=1001
		inner join (select Valor from fnc_getTblValoresTexto (@sAnalistas)) AnaL on (AnaL.Valor=RH.cPersCod)
	where  datediff(day,CC.dvenc,CC.dPago)<=@ndiaMora and CC.dPago between @dfechaIni and @dfechaFin and 
		  CC.nNroCalen=(select max(nNroCalen) from colocCalendario where cCtaCod=CC.cCtaCod)
		and Substring(P.cCtaCod,4,2) in (select Valor from fnc_getTblValoresTexto (@sAgencias)) 
	order by substring(P.cCtaCod,9,1),substring(C.cLineaCred,6,1),substring(P.cCtaCod,6,3),CL.cDescripcion
end
