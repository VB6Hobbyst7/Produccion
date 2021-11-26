set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go

ALTER  procedure [dbo].[stp_sel_ReportePDTJornadaLab]
@dfechaIni varchar(8),
@dfechaFin varchar(8)
as

select case len(PE.cPersIDTpo)
		when 1 then '0'+ convert(varchar(1),PE.cPersIDTpo)
		else convert(varchar(1),PE.cPersIDTpo)
		end TpoDoc,PE.cPersIDNro cPersIDnro,
		RHPDC.cPlanillaCod,
		cRHConceptoDescripcion,		
		PE.cPersCod,
		convert(decimal(12,0),IsNull(Round(30 * sum(RHPDC.nMonto),0),0)) as DiasTraba,
		convert(decimal(12,0),IsNull(Round(30 * sum(RHPDC.nMonto),0),0)*8) as HorasTraba,
		0 as MinutosTraba,
		convert(decimal(12,0),IsNull(Round(sum(D1.horasExtras),0),0)) as HorasExtras,
		0 as MinutosExtras
from RHEMPLEADO RHE inner join PersID PE on (RHE.cPersCod=PE.cPersCod)
		inner join RHPlanillaDetCon RHPDC on (PE.cPersCod=RHPDC.cPersCod)
		left join rhConceptotabla RHCT on (RHPDC.cRHConceptoCod=RHCT.cRHConceptoCod)
		left join	(
					Select Sum(nMonto) horasExtras, cPersCod 
						FROM RHPlanillaDetCon 
						Where	cRRHHPeriodo Between @dfechaIni And @dfechaFin 
						And cRHConceptoCod In (121) --And cPlanillaCod In ('P01','E01')            
						and cRRHHPeriodo Between @dfechaIni And @dfechaFin 
						Group By cPersCod
					) D1 On D1.cPersCod = PE.cPersCod 
Where cRRHHPeriodo Between @dfechaIni And @dfechaFin 
	And RHPDC.cPlanillaCod In ('P01','E01')  
	
	And RHCT.cRHConceptoCod In (404)
	and PE.cPersIDTpo=1
group by PE.cPersIDTpo,PE.cPersIDNro,RHPDC.cPlanillaCod,cRHConceptoDescripcion,PE.cPersCod
having Sum(RHPDC.nMonto)>0
order by RHPDC.cPlanillaCod,PE.cPersCod
