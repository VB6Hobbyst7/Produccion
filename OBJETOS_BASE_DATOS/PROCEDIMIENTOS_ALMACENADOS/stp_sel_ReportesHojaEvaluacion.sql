create procedure stp_sel_ReportesHojaEvaluacion
(
@cNumFuente varchar(8),
@dFecha datetime
)
as
begin
	select	isnull(PHE.dPersEval,'19000101') as dPersEval,isnull(PHE.cNumFuente,'00000000') as cNumFuente,HE.cCodHojEval,HE.cDescripcion,
			isnull(PHE.nPersonal,0.00) as nPersonal,isnull(PHE.nNegocio,0.00) as
			nNegocio,isnull(PHE.nUnico,0.00) as nUnico
	from HojaEvaluacion HE left join  
				 (select nEstado,dPersEval,cNumFuente,cCodHojEval,
						sum(case when nTipoImporte=1 then nImporte else 0 end) as nPersonal,
						sum(case when nTipoImporte=2 then nImporte else 0 end) as nNegocio,
						sum(case when nTipoImporte=3 then nImporte else 0 end) as nUnico
				   from PersFIHojaEvaluacion
				   where cNumFuente=@cNumFuente
						and datediff(day,dPersEval,@dFecha)=0
						and nEstado=1
				   group by cCodHojEval,dPersEval,cNumFuente,nEstado) PHE
				   on HE.cCodHojEval=PHE.cCodHojEval
	where HE.cCodHojEval not like '4%'
	order by HE.cCodHojEval
end