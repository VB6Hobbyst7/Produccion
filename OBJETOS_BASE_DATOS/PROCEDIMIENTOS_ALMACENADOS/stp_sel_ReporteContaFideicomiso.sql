create procedure stp_sel_ReporteContaFideicomiso
(
@dFechaI datetime,
@ntipoCam decimal(10,2)
)
as
begin
	select isnull(' ' + A.cAgeCod +' '+ A.cAgeDescripcion,'Total') 
	descripcion ,
	sum(x.nSMoratotio)nSMoratotio ,sum(x.nSCompensatorio) nSCompensatorio,
	 sum(x.nSInteresGrac) nSInteresGrac ,sum(x.nSMoratotio + x.nSCompensatorio + x.nSInteresGrac) as nSubTo
	from agencias A inner join 
	(Select 	SubString(C.cCtaCod,4,2) codAge,
			SubString(C.cCtaCod,9,1) nMoneda,
			case SubString(C.cCtaCod,9,1)
				when 1 then	Sum(case MCD.nPrdConceptoCod when 1101 then MCD.nMonto else 0.0 end)
				else Sum(case MCD.nPrdConceptoCod when 1101 then MCD.nMonto else 0.0 end)*@ntipoCam end nSMoratotio,
			case SubString(C.cCtaCod,9,1)
				when 1 then	Sum(case MCD.nPrdConceptoCod when 1100 then MCD.nMonto else 0.0 end)
				else Sum(case MCD.nPrdConceptoCod when 1100 then MCD.nMonto else 0.0 end)*@ntipoCam end	 nSCompensatorio,
			case SubString(C.cCtaCod,9,1)
				when 1 then	Sum(case MCD.nPrdConceptoCod when 1102 then MCD.nMonto else 0.0 end)
				else Sum(case MCD.nPrdConceptoCod when 1102 then MCD.nMonto else 0.0 end)*@ntipoCam end	 nSInteresGrac			 
			
	From Mov M Inner Join MovColDet MCD On M.nMovNro = MCD.nMovNro
		 And M.nMovFlag = 0 And substring(M.cMovNro,1,4) = year(@dFechaI)  And substring(M.cMovNro,5,2) = month(@dFechaI)
		 And (MCD.cOpeCod Like '100[234567]%' Or MCD.cOpeCod = '102100')
		 Inner Join Colocaciones C On C.cCtaCod = MCD.cCtaCod 
		 --inner join DBConsolidada..colocLineaCreditoEquiv clce on (C.cLineaCred=clce.cLineaCred) and clce.cCtaCont='02'
	Where C.cLineaCred Like '0401%' And MCD.nPrdConceptoCod In (1101,1100,1102)
	Group By SubString(C.cCtaCod,4,2),SubString(C.cCtaCod,9,1)
	--order by SubString(C.cCtaCod,4,2) asc
	)X on X.codAge=A.cAgeCod
	group by ' ' + A.cAgeCod +' '+ A.cAgeDescripcion with rollup
	order by isnull(' ' + A.cAgeCod +' '+ A.cAgeDescripcion,'Total')
end