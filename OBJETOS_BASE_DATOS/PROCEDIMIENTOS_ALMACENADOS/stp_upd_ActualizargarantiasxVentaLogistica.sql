create procedure stp_upd_ActualizargarantiasxVentaLogistica
@cNumGarant char(8),
@nEstadoAdju int,
@dEstadoAdju datetime,
@cUsuariAdju varchar(25),
@dFechaCompra datetime,
@nMonedaAdju int,
@nTipoCambio decimal(18,4),
@cPersCodComprador varchar(13),
@nVendido int
as 
update garantias set 
	nEstadoAdju=@nEstadoAdju,
	dEstadoAdju=@dEstadoAdju,
	cUsuariAdju=@cUsuariAdju,
	dFechaCompra=@dFechaCompra,
	nMonedaAdju=@nMonedaAdju,
	nTipoCambio=@nTipoCambio,
	cPersCodComprador=@cPersCodComprador,
	nVendido=@nVendido
where cNumGarant=@cNumGarant