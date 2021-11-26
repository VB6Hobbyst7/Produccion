ALTER  procedure [dbo].[stp_sel_ReportePDTEstablecimiento]
@dfechaIni varchar(8),
@dfechaFin varchar(8)
as
select case len(PE.cPersIDTpo)
		when 1 then '0'+ convert(varchar(1),PE.cPersIDTpo)
		else convert(varchar(1),PE.cPersIDTpo)
		end TpoDoc,PE.cPersIDNro cPersIDnro,
		RHPDC.cPlanillaCod,		
		PE.cPersCod,
		RRHH.cAgenciaActual cAgeCod,
		'20103845328' as sRUCEmpresa,
		Ag.cAgeDescripcion,
		0 ntasa
from	RRHH RRHH inner join RHEMPLEADO RHE on (RRHH.cPersCod=RHE.cPersCod)
		inner join PersID PE on (RHE.cPersCod=PE.cPersCod)
		inner join RHPlanillaDetCon RHPDC on (PE.cPersCod=RHPDC.cPersCod)
		left join rhConceptotabla RHCT on (RHPDC.cRHConceptoCod=RHCT.cRHConceptoCod)		
		left join agencias Ag on (substring(PE.cPersCod,4,2)=Ag.cAgeCod)
Where cRRHHPeriodo Between @dfechaIni And @dfechaFin 
	And RHPDC.cPlanillaCod In ('P01','E01')  
	and PE.cPersIDTpo=1
group by PE.cPersIDTpo,PE.cPersIDNro,RHPDC.cPlanillaCod,PE.cPersCod,RRHH.cAgenciaActual,Ag.cAgeDescripcion
having Sum(RHPDC.nMonto)>0
order by RHPDC.cPlanillaCod,PE.cPersCod