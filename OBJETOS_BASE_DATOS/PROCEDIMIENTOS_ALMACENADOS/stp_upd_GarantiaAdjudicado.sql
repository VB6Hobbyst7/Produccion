alter procedure stp_upd_GarantiaAdjudicado
@cNumGarant char(8),
@dEstadoA datetime,
@nEstadoA int,
@cUsuariA varchar(25)

as

declare @cUsuario varchar(25),@nEstado int,@nEstadoAdj int,@dFechaAdj datetime

select	@cUsuario=isnull(cUsuariAdju,''),
		@nEstado=nEstado,
		@nEstadoAdj=isnull(nEstadoAdju,0),
		@dFechaAdj=isnull(dEstadoAdju,'1990/01/01') 
from garantias where cNumGarant=@cNumGarant
	
insert into garantiaEstado (cNumGarant,nEstado,nEstadoAdju,dEstadoAdju,cUsuariAdju)
values (@cNumGarant,@nEstado,@nEstadoAdj,@dFechaAdj,@cUsuario)			
	
update garantias set
nEstadoAdju =@nEstadoA,
dEstadoAdju =@dEstadoA,
cUsuariAdju =@cUsuariA
where cNumGarant=@cNumGarant