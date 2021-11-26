create procedure stp_sel_reporteGarantiasLogisticaxPeriodo
@dFecSis datetime
as

--declare @dFecSis datetime
--set @dFecSis='2008/08/11'

SELECT X.cPersCod,
	   X.cNombreTitular,
	   X.cCtaCod,
	   X.cNumGarant,
	   X.cDescripcion,
	   X.cDireccion,
	   X.nMontSan,
	   X.nMontoRetasacion as nValorComerc,
	   X.nMontoRetasacion*0.8 as nValorRealizacion,
	   convert(varchar(12),convert(decimal(18,2),X.nMontoFrente)) + ' Ml' as cFrente,
	   convert(varchar(12),convert(decimal(18,2),X.nMontoDerech)) + ' Ml' as cDerech,
	   convert(varchar(12),convert(decimal(18,2),X.nMontoIzquie)) + ' Ml' as cIzquie,
	   convert(varchar(12),convert(decimal(18,2),X.nMontoFondoV)) + ' Ml' as cFondo,
	   X.nEstadoAdju
FROM (
select	PPer.cPersCod,
		(select cPersNombre from Persona where cPersCod=PPer.cPersCod) as cNombreTitular,
		ColG.cCtaCod,
		Gara.cNumGarant,
		Gara.cDescripcion,
		Gara.cDireccion,
		isnull((select 
			case nMoneda
				when 0 then sum(nMontSan) 
				else 
			sum(nMontSan) * isnull((select nValVent from tipoCambio where datediff(d,dFecCamb,dFecSan)=0),0)
			end 			
			from GarantiaSaneamiento where cNumGarant=Gara.cNumGarant and nPeriSan=year(@dFecSis) and nTESan=2 and nEstado=1 group by nMoneda,dFecSan),0)
			nMontSan,
			isnull((select sum(nValor) from garantiaLogistica where cNumGarant=Gara.cNumGarant and nPeriodo=year(@dFecSis)	and nTValor like '1%' and nEstado=1),0) as nMontoRetasacion,
			isnull((select sum(nValor) from garantiaLogistica where cNumGarant=Gara.cNumGarant and nPeriodo=year(@dFecSis)	and nTValor = 20 and nEstado=1),0) as nMontoFrente,
			isnull((select sum(nValor) from garantiaLogistica where cNumGarant=Gara.cNumGarant and nPeriodo=year(@dFecSis)	and nTValor = 21 and nEstado=1),0) as nMontoDerech,
			isnull((select sum(nValor) from garantiaLogistica where cNumGarant=Gara.cNumGarant and nPeriodo=year(@dFecSis)	and nTValor = 22 and nEstado=1),0) as nMontoIzquie,
			isnull((select sum(nValor) from garantiaLogistica where cNumGarant=Gara.cNumGarant and nPeriodo=year(@dFecSis)	and nTValor = 23 and nEstado=1),0) as nMontoFondoV,
			isnull(nEstadoAdju,0) nEstadoAdju
from Garantias Gara 
		Inner Join ColocGarantia ColG on (Gara.cNumGarant=ColG.cNumGarant)
		Inner Join ProductoPersona PPer on ColG.cCtaCod=PPer.cCtaCod and PPer.nPrdPersRelac=20		
where nEstadoAdju in (8,10) 
) X
/*
select sum(nValor) 
from garantiaLogistica where cNumGarant=('00071507') and nPeriodo=year(@dFecSis)
and nTValor like '1%'
exec stp_sel_reporteGarantiasLogisticaxPeriodo '2008/08/21'
*/