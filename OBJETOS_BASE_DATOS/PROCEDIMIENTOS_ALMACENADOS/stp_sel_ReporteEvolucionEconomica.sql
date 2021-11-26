alter procedure  stp_sel_ReporteEvolucionEconomica
(
	@cPersCod char(13)
)
as
begin
	select * into #Temp from
	(select 1 as nOrden,'ACTIVO CORRIENTE' as cDescrip, '1001' as cCod, 1 as SumaResta union 
	 select 1 as nOrden,'ACTIVO CORRIENTE' as cDescrip, '1002' as cCod, 1 as SumaResta union
	 select 2 as nOrden,'INVENTARIO' as cDescrip, '1002' as cCod, 1 as SumaResta union
	 select 3 as nOrden,'ACTIVO FIJO' as cDescrip, '1003' as cCod, 1 as SumaResta union
	 select 4 as nOrden,'ACTIVO TOTAL' as cDescrip, '1001' as cCod, 1 as SumaResta union
	 select 4 as nOrden,'ACTIVO TOTAL' as cDescrip, '1002' as cCod, 1 as SumaResta union
	 select 4 as nOrden,'ACTIVO TOTAL' as cDescrip, '1003' as cCod, 1 as SumaResta union
	 select 5 as nOrden,'PASIVO CORRIENTE' as cDescrip, '1004' as cCod, 1 as SumaResta union
	 select 6 as nOrden,'PASIVO NO CORRIENTE' as cDescrip, '1005' as cCod, 1 as SumaResta union
	 select 7 as nOrden,'PASIVO TOTAL' as cDescrip, '1004' as cCod, 1 as SumaResta union
	 select 7 as nOrden,'PASIVO TOTAL' as cDescrip, '1005' as cCod, 1 as SumaResta union
	 select 8 as nOrden,'PATRIMONIO' as cDescrip, '1001' as cCod, 1 as SumaResta union
	 select 8 as nOrden,'PATRIMONIO' as cDescrip, '1002' as cCod, 1 as SumaResta union
	 select 8 as nOrden,'PATRIMONIO' as cDescrip, '1003' as cCod, 1 as SumaResta union
	 select 8 as nOrden,'PATRIMONIO' as cDescrip, '1004' as cCod, -1 as SumaResta union
	 select 8 as nOrden,'PATRIMONIO' as cDescrip, '1005' as cCod, -1 as SumaResta union
	 select 9 as nOrden,'TOTAL PASIVO Y PATRIMONIO' as cDescrip, '1004' as cCod, 1 as SumaResta union
	 select 9 as nOrden,'TOTAL PASIVO Y PATRIMONIO' as cDescrip, '1005' as cCod, 1 as SumaResta union
	 select 9 as nOrden,'TOTAL PASIVO Y PATRIMONIO' as cDescrip, '1001' as cCod, 1 as SumaResta union
	 select 9 as nOrden,'TOTAL PASIVO Y PATRIMONIO' as cDescrip, '1002' as cCod, 1 as SumaResta union
	 select 9 as nOrden,'TOTAL PASIVO Y PATRIMONIO' as cDescrip, '1003' as cCod, 1 as SumaResta union
	 select 9 as nOrden,'TOTAL PASIVO Y PATRIMONIO' as cDescrip, '1004' as cCod, -1 as SumaResta union
	 select 9 as nOrden,'TOTAL PASIVO Y PATRIMONIO' as cDescrip, '1005' as cCod, -1 as SumaResta UNION
	 select 10 as nOrden,'INGRESOS' as cDescrip, '2001' as cCod, 1 as SumaResta union
	 select 11 as nOrden,'COSTOS VENTAS' as cDescrip, '2002' as cCod, 1 as SumaResta union
	 select 12 as nOrden,'OTROS EGRESOS' as cDescrip, '2003' as cCod, 1 as SumaResta union
	 select 13 as nOrden,'INGRESO NETO' as cDescrip, '2001' as cCod, 1 as SumaResta union
	 select 13 as nOrden,'INGRESO NETO' as cDescrip, '2002' as cCod, -1 as SumaResta union
	 select 13 as nOrden,'INGRESO NETO' as cDescrip, '2003' as cCod, -1 as SumaResta 
	) T

	insert into TablaRepo(cDescripcion) values ('FECHA')
	insert into TablaRepo(cDescripcion) values ('ACTIVO CORRIENTE')
	insert into TablaRepo(cDescripcion) values ('INVENTARIO')
	insert into TablaRepo(cDescripcion) values ('ACTIVO FIJO')
	insert into TablaRepo(cDescripcion) values ('ACTIVO TOTAL')
	insert into TablaRepo(cDescripcion) values ('PASIVO CORRIENTE')
	insert into TablaRepo(cDescripcion) values ('PASIVO NO CORRIENTE')
	insert into TablaRepo(cDescripcion) values ('PASIVO TOTAL')
	insert into TablaRepo(cDescripcion) values ('PATRIMONIO')
	insert into TablaRepo(cDescripcion) values ('TOTAL PASIVO Y PATRIMONIO')
	insert into TablaRepo(cDescripcion) values ('INGRESOS')
	insert into TablaRepo(cDescripcion) values ('COSTOS VENTAS')
	insert into TablaRepo(cDescripcion) values ('OTROS EGRESOS')
	insert into TablaRepo(cDescripcion) values ('INGRESO NETO')

	declare @cNumFuente char(8)
	declare @cRazSocDescrip varchar(50)
	declare @dPersEval datetime
	declare @Cont int

	declare curFue cursor for 
		select DISTINCT top 5 pfi.cnumfuente,PFI.cRazSocDescrip,PHE.dPersEval
		from persfteingreso pfi
		inner join colocfteingreso cfi on pfi.cnumfuente=cfi.cnumfuente
		inner join producto p on cfi.cctacod=p.cctacod and p.nprdestado not in (2003,2040,2080,2090,2091,2092)
		inner join productopersona pp on p.cctacod=pp.cctacod and pp.nprdpersrelac=28
		inner join rrhh rh on pp.cperscod=rh.cperscod
		inner join constante con on p.nprdEstado=con.nconsvalor and nconscod= 3001
		inner join colocaciones c on p.cctacod=c.cctacod
		inner join colocacestado ce on p.cctacod=ce.cctacod and ce.nprdestado=2000
		INNER JOIN persfihojaevaluacion PHE ON PHE.nEstado=1 AND PFI.cnumfuente=PHE.cnumfuente
		where pfi.cperscod=@cPersCod
		order by PHE.dPersEval

	open curFue
	fetch next from curFue into @cNumFuente, @cRazSocDescrip, @dPersEval

	set @Cont=1
	while (@@fetch_status=0)
	begin
		if (@Cont=1)
		begin
			update TablaRepo set cFecha1=convert(char(10),@dPersEval,103) where cDescripcion='FECHA'
			update TablaRepo set cFecha1=F.nImporte from TablaRepo T2 inner join 
				(select T.nOrden,T.cDescrip, ISNULL(sum(PH.nImporte*T.SumaResta),0) as nImporte
				from #Temp T inner join PersFIHojaEvaluacion PH on T.cCod=left(PH.cCodHojEval,4)
				where PH.cNumFuente=@cNumFuente and datediff(day,PH.dPersEval,@dPersEval)=0
				group by T.nOrden,T.cDescrip) F on T2.cDescripcion=F.cDescrip			
				update TablaRepo set cFecha1=0 where cFecha1 is null
		end
		if (@Cont=2)
		begin
			update TablaRepo set cFecha2=convert(char(10),@dPersEval,103) where cDescripcion='FECHA'
			update TablaRepo set cFecha2=F.nImporte from TablaRepo T2 inner join 
				(select T.nOrden,T.cDescrip, sum(PH.nImporte*T.SumaResta) as nImporte
				from #Temp T inner join PersFIHojaEvaluacion PH on T.cCod=left(PH.cCodHojEval,4)
				where PH.cNumFuente=@cNumFuente and datediff(day,PH.dPersEval,@dPersEval)=0
				group by T.nOrden,T.cDescrip) F on T2.cDescripcion=F.cDescrip
				update TablaRepo set cFecha2=0 where cFecha2 is null

		end
		if (@Cont=3)
		begin
			update TablaRepo set cFecha3=convert(char(10),@dPersEval,103) where cDescripcion='FECHA'
			update TablaRepo set cFecha3=F.nImporte from TablaRepo T2 inner join 
				(select T.nOrden,T.cDescrip, sum(PH.nImporte*T.SumaResta) as nImporte
				from #Temp T inner join PersFIHojaEvaluacion PH on T.cCod=left(PH.cCodHojEval,4)
				where PH.cNumFuente=@cNumFuente and datediff(day,PH.dPersEval,@dPersEval)=0
				group by T.nOrden,T.cDescrip) F on T2.cDescripcion=F.cDescrip
				update TablaRepo set cFecha3=0 where cFecha3 is null
		end
		if (@Cont=4)
		begin
			update TablaRepo set cFecha4=convert(char(10),@dPersEval,103) where cDescripcion='FECHA'
			update TablaRepo set cFecha4=F.nImporte from TablaRepo T2 inner join 
				(select T.nOrden,T.cDescrip, sum(PH.nImporte*T.SumaResta) as nImporte
				from #Temp T inner join PersFIHojaEvaluacion PH on T.cCod=left(PH.cCodHojEval,4)
				where PH.cNumFuente=@cNumFuente and datediff(day,PH.dPersEval,@dPersEval)=0
				group by T.nOrden,T.cDescrip) F on T2.cDescripcion=F.cDescrip
				update TablaRepo set cFecha4=0 where cFecha4 is null
		end
		if (@Cont=5)
		begin
			update TablaRepo set cFecha5=convert(char(10),@dPersEval,103) where cDescripcion='FECHA'
			update TablaRepo set cFecha5=F.nImporte from TablaRepo T2 inner join 
				(select T.nOrden,T.cDescrip, sum(PH.nImporte*T.SumaResta) as nImporte
				from #Temp T inner join PersFIHojaEvaluacion PH on T.cCod=left(PH.cCodHojEval,4)
				where PH.cNumFuente=@cNumFuente and datediff(day,PH.dPersEval,@dPersEval)=0
				group by T.nOrden,T.cDescrip) F on T2.cDescripcion=F.cDescrip
				update TablaRepo set cFecha5=0 where cFecha5 is null
		end

		set @Cont=@Cont+1
		fetch next from curFue into @cNumFuente, @cRazSocDescrip, @dPersEval
	end
	close curFue
	deallocate curFue	
end
