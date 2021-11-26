create procedure stp_sel_ReporteClientesConDeudaCanceladaNoVueltosAtender
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
	isnull((select top 1
		case nColocCalendCod 
			when 40 then 'FecFija/30' 
			when 41 then 'FecFija/30'
			when 50 then 'FecFija/30'
			when 51 then 'FecFija/30'
			when 60 then 'FecFija/30'
			when 61 then 'FecFija/30'
			else 'PerFijo/' + convert(varchar(10),nPlazo) 
		end X
	from colocacestado where cCtaCod = P.cCtaCod),'Sin') cPlazo2,

	P.cCtaCod,PE.cPersNombre,C.nMontoCol,
	cCalifActual= isnull(case when CP.cCalGen='0' then '0 NORMAL'        when CP.cCalGen='1' then '1 CPP'        when CP.cCalGen='2' then '2 DEFICIENTE'       when CP.cCalGen='3' then '3 DUDOSO'       when CP.cCalGen='4' then '4 PERDIDA' end,'SIN CALI') ,
	PE.cPersDireccDomicilio,
	PE.cPersTelefono,
	NroCuotaApr=(Select Count(*) From ColocCalendario CD Where CD.cCtaCod=P.cCtaCod and CD.nColocCalendApl=1 and CD.nNroCalen=CC.nNroCalen),
	P.dPrdEstado,RH.cUser


	from producto P inner join Colocaciones C on P.cCtaCod=C.cCtaCod and len(C.cLineaCred)=13 and P.nPrdEstado=2050
		inner join ColocLineaCredito CL on (C.cLineaCred=CL.cLineaCred) and CL.bEstado=1
		inner join Productopersona PP on (P.cCtaCod=PP.cCtaCod) and PP.nPrdPersRelac=20
		inner join Persona PE on (PP.cPersCod=PE.cPersCod)
		inner join colocCalendario CC on (CC.cCtaCod=C.cCtaCod) and CC.nColocCalendApl=1 and CC.nColocCalendEstado=1		
		inner join ProductoPersona PP2 on (P.cCtaCod=PP2.cCtaCod) and PP2.nPrdPersRelac=28
		inner join RRHH RH on (PP2.cPersCod=RH.cPersCod)
		inner join constante Cons1 on (Cons1.nConsValor=substring(P.cCtaCod,6,3)) and nConsCod=1001
		left join ColocCalifProv CP on CP.cCtaCod = P.cCtaCod 
		inner join (select Valor from fnc_getTblValoresTexto (@sAnalistas)) AnaL on (AnaL.Valor=RH.cPersCod)

	where   datediff(day,CC.dvenc,CC.dPago)<=@ndiaMora and P.dPrdEstado between @dfechaIni and @dfechaFin and 
		  CC.nNroCalen=(select max(nNroCalen) from colocCalendario where cCtaCod=CC.cCtaCod)	
		and Substring(P.cCtaCod,4,2) in (select Valor from fnc_getTblValoresTexto (@sAgencias)) 
		and PP.cPersCod not in (select PP3.cPersCod from productopersona PP3 inner join colocaciones col2 on
						PP3.nPrdPersRelac=20 and (PP.cCtaCod<>PP3.cCtaCod and PP3.cCtaCod=col2.cCtaCod) 
						where col2.dVigencia>P.dPrdEstado)
	group by C.cLineaCred,CL.cDescripcion,Cons1.cConsDescripcion,P.cCtaCod,PE.cPersNombre,C.nMontoCol,P.nSaldo,
	CC.nNroCalen,CC.nColocCalendApl,
	P.dPrdEstado,RH.cUser,CP.cCalGen,PE.cPersDireccDomicilio,PE.cPersTelefono
	order by substring(P.cCtaCod,9,1),substring(C.cLineaCred,6,1),substring(P.cCtaCod,6,3),CL.cDescripcion
end