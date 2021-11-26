Create Procedure stp_sel_ValidarRevision
(
@CodCta as varchar(50),
@FCierre as char (10)
)
as
Select * from RevisionCalificacion
where vCodCta=@CodCta and cFCierre=@FCierre and iEstado=1