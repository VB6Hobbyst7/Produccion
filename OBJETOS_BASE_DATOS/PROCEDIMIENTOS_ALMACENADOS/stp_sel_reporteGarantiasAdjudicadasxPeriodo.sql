alter procedure stp_sel_reporteGarantiasAdjudicadasxPeriodo
@dFecSis datetime
as

--declare @dFecSis datetime
--set @dFecSis='2008/08/11'

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
			from GarantiaSaneamiento where cNumGarant=Gara.cNumGarant and nEstado=1 and nPeriSan=year(@dFecSis) and nTESan=1 group by nMoneda,dFecSan),0)
			nMontSan,
			isnull((select cPersCod from GarantiaRemate where cNumGarant=Gara.cNumGarant and (cPersCod!='')),'') as Comprador,
			isnull((select cPersNombre from Persona where cPersCod=isnull((select cPersCod from GarantiaRemate where cNumGarant=Gara.cNumGarant and (cPersCod!='')),'')),'') as cDComprador,			
				isnull((select 
				case nMoneda
					when 0 then sum(nMonto) 
					else 
						sum(nMonto) * isnull((select nValVent from tipoCambio where datediff(d,dFecCamb,dFechaRemate)=0),0)
				end 					
				from GarantiaRemate where cNumGarant=Gara.cNumGarant and (cPersCod!='') group by nMoneda,dFechaRemate),0) as MontoCompra,
			isnull(nEstadoAdju,0) nEstadoAdju
from Garantias Gara 
		Inner Join ColocGarantia ColG on (Gara.cNumGarant=ColG.cNumGarant)
		Inner Join ProductoPersona PPer on ColG.cCtaCod=PPer.cCtaCod and PPer.nPrdPersRelac=20		
where nEstadoAdju in (7,8,9,10) 

--update garantias set nEstadoAdju =8 where cNumGarant ='00083480'
--1090200016322