Create Procedure stp_sel_UsuarioCMACMAYNAS
as
select vNombre
, vUsuario
, vAgencia
, vArea
, vGrupo
, tOperaciones
, tColocaciones
, tOtros
from UsuarioCMACMAYNASTem
Go