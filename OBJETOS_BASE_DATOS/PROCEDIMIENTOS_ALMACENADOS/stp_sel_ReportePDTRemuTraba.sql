create  procedure [dbo].[stp_sel_ReportePDTRemuTraba]
@dfechaIni varchar(8),
@dfechaFin varchar(8)
as
select 
		TpoDoc,
		cPersIDnro,
		cRHConceptoCodSUNAT, 
		cPersona,		
		Sum(MontoPagado) as MontoPagado,
		Sum(MontoPagado) as MontoDeven
from
(
		select	
		'01' TpoDoc,
		case PE.cPersIDTpo
			when 1 then PE.cPersIDNro 
			else substring(PE.cPersIDNro,3,8)
		end	cPersIDnro,
		isnull(cRHConceptoCodSUNAT,'') cRHConceptoCodSUNAT, 
		PE.cPersCod cPersona,
		Sum(RHPDC.nMonto) as MontoPagado,
		Sum(RHPDC.nMonto) as MontoDeven
from RHEMPLEADO RHE inner join PersID PE on (RHE.cPersCod=PE.cPersCod)
		inner join RHPlanillaDetCon RHPDC on (PE.cPersCod=RHPDC.cPersCod)
		inner join rhConceptotabla RHCT on (RHPDC.cRHConceptoCod=RHCT.cRHConceptoCod)
		inner join rhConceptotablasunat RHCTS on (RHCT.cRHConceptoCod=RHCTS.cRHConceptoCod)
Where cRRHHPeriodo Between @dfechaIni And @dfechaFin 

	And RHPDC.cPlanillaCod In ('E01','E02','E04','E05','E06','E08')  
	
	and RHCT.cRHConceptoCod not in (215,112,255,260,261,303,404,422,112,163,215,417,130,413,421,240,248)
	
	and PE.cPersIDTpo=1
group by 
		PE.cPersIDTpo,PE.cPersIDNro,PE.cPersCod,cRHConceptoCodSUNAT
having Sum(RHPDC.nMonto)>0
) X1
group by

		TpoDoc,
		cPersIDnro,
		cRHConceptoCodSUNAT, 
		cPersona

order by cPersona,cRHConceptoCodSUNAT