Create Procedure stp_sel_ValidarCierre
(
@FCierre as varchar(50)
)
as 
select top 10 *
from DBConsolidada..ColocCalifProvtotal
where convert(varchar(10), dFecha, 112)=@FCierre
order by dFecha