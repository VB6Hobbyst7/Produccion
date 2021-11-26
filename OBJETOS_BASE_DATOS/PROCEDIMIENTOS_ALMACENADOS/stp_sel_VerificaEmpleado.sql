create procedure stp_sel_VerificaEmpleado
(
@CodCli char(13)
)
as
begin
	select Count(cPersCod) as NroEmp  from 
	(Select cperscod From RRHH Where cPerscod =@CodCli And nRHEstado = 201 
	union
	Select cperscod From persona Where cPerscod =@CodCli And nPersRelaInst=1) xy
end